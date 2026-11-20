# 多账号密钥扫描器 —— 基于仓库 decrypt/windows/find_key.ps1 改造
# 差异：仓库原版只拿"最大的 message.db"当校验靶、且找到 1 把就停；
#       本机有 8 个账号且各有独立 key，故这里一次内存扫描同时匹配全部账号的库头。
# 只读内存与库文件；结果写入 decrypt/windows/wxwork_keys.json (已被 .gitignore 忽略)
# 用法: powershell -NoProfile -ExecutionPolicy Bypass -File find_all_keys.ps1 [-Wide]
#       -Wide 会额外扫描非私有/只读/可执行内存页（慢，用于找历史登录账号的 key）
param([switch]$Wide)
# -File 方式传开关在部分宿主会被吞掉，额外支持环境变量
if ($env:WECOM_WIDE_SCAN -eq '1') { $Wide = $true }
$ErrorActionPreference = 'Continue'
# PowerShell 5.1 默认按 OEM/GBK 输出到管道，Python 侧按 UTF-8 解码会乱码
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

$cs = @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

public class MultiKeyScan {
    [DllImport("kernel32.dll", SetLastError=true)] static extern IntPtr OpenProcess(uint a, bool inh, int pid);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
    [DllImport("kernel32.dll", SetLastError=true)] static extern long VirtualQueryEx(IntPtr h, IntPtr addr, out MBI mbi, uint len);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool ReadProcessMemory(IntPtr h, IntPtr addr, byte[] buf, IntPtr size, out IntPtr read);
    [DllImport("advapi32.dll", SetLastError=true)] static extern bool OpenProcessToken(IntPtr h, uint acc, out IntPtr tok);
    [DllImport("advapi32.dll", SetLastError=true)] static extern bool LookupPrivilegeValue(string s, string n, out long luid);
    [DllImport("advapi32.dll", SetLastError=true)] static extern bool AdjustTokenPrivileges(IntPtr tok, bool dis, ref TP newp, uint len, IntPtr prev, IntPtr ret);
    [DllImport("kernel32.dll")] static extern IntPtr GetCurrentProcess();

    [StructLayout(LayoutKind.Sequential)] struct MBI {
        public IntPtr BaseAddress, AllocationBase;
        public uint AllocationProtect, __a1;
        public IntPtr RegionSize;
        public uint State, Protect, Type, __a2;
    }
    [StructLayout(LayoutKind.Sequential)] struct TP { public uint Count; public long Luid; public uint Attr; }

    const uint MEM_COMMIT=0x1000, MEM_PRIVATE=0x20000;
    const uint PAGE_RW=0x04, PAGE_WC=0x08, PAGE_GUARD=0x100, PAGE_NOACCESS=0x01;

    static byte[][] FRAGS, CIPHERS, WANT;
    static bool[] matched;
    static byte[] IV, SUF;
    static bool WIDE = false;
    static MD5 md5 = MD5.Create();
    static Aes AES;

    public static void SetWide(bool w) { WIDE = w; }

    public static void EnableDebug() {
        IntPtr tok; if (!OpenProcessToken(GetCurrentProcess(), 0x20|0x08, out tok)) return;
        long luid; if (!LookupPrivilegeValue(null, "SeDebugPrivilege", out luid)) return;
        TP tp = new TP(); tp.Count=1; tp.Luid=luid; tp.Attr=0x2;
        AdjustTokenPrivileges(tok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
    }

    static int PC(ulong x){ int c=0; while(x!=0){ x&=x-1; c++; } return c; }

    // 廉价筛选与账号靶无关：每个 16 字节窗口只算一次，不再按 8 个靶重复 8 遍
    static bool Candidate(byte[] b, int o) {
        ulong m0=0,m1=0,m2=0,m3=0; int na=0;
        for (int j=0;j<16;j++){
            byte v=b[o+j];
            int w=v>>6; ulong bit=1UL<<(v&63);
            if(w==0)m0|=bit; else if(w==1)m1|=bit; else if(w==2)m2|=bit; else m3|=bit;
            if(v<0x20||v>0x7e) na++;
        }
        if (na<3) return false;
        if (PC(m0)+PC(m1)+PC(m2)+PC(m3) < 11) return false;
        return true;
    }

    // 通过筛选后：MD5+AES.Key 只算一次，再逐个靶做一次 16 字节解密校验
    static List<string> Verify(byte[] b, int o) {
        var hits = new List<string>();
        byte[] inp=new byte[24];
        Array.Copy(b,o,inp,0,16); Array.Copy(SUF,0,inp,16,8);
        byte[] pk=md5.ComputeHash(inp);
        AES.Key=pk;
        string hex = null;
        // 同一个候选 key 只建一次解密器，8 个靶复用（TransformFinalBlock 会自动复位）
        using (var dec=AES.CreateDecryptor()) {
            for (int t=0; t<FRAGS.Length; t++) {
                if (matched[t]) continue;
                byte[] pt=dec.TransformFinalBlock(CIPHERS[t],0,16);
                bool ok=true;
                for(int j=0;j<8;j++) if(pt[j]!=WANT[t][j]) { ok=false; break; }
                if(!ok) continue;
                if (hex==null) {
                    var sb=new StringBuilder();
                    for(int j=0;j<16;j++) sb.Append(b[o+j].ToString("x2"));
                    hex=sb.ToString();
                }
                matched[t]=true;
                hits.Add(hex + "|" + t);
            }
        }
        return hits;
    }

    static bool AllDone(){ foreach(bool d in matched) if(!d) return false; return true; }

    // 本机有 8 个账号但内存里通常只有 2–3 把 key，AllDone 永远不会成立，
    // 没有时间预算就会把整个进程内存扫完（实测 >10 分钟）。给一个硬上限保证界面可用。
    public static int MaxSeconds = 120;
    public static bool Limited = false;
    public static double ElapsedSeconds = 0;
    public static int PreMatched = 0;
    public static void SetMaxSeconds(int s) { if (s > 0) MaxSeconds = s; }

    // 历史已抓到的 key 先验一遍并标记，扫描时不再为这些账号重复做 AES 校验
    static int MarkKnown(string[] known) {
        int n=0;
        if (known==null) return 0;
        foreach (string hk in known) {
            if (string.IsNullOrEmpty(hk) || hk.Length!=32) continue;
            byte[] k=new byte[16];
            try { for(int i=0;i<16;i++) k[i]=Convert.ToByte(hk.Substring(i*2,2),16); } catch { continue; }
            byte[] inp=new byte[24];
            Array.Copy(k,0,inp,0,16); Array.Copy(SUF,0,inp,16,8);
            AES.Key=md5.ComputeHash(inp);
            using (var dec=AES.CreateDecryptor()) {
                for (int t=0;t<FRAGS.Length;t++) {
                    if (matched[t]) continue;
                    byte[] pt=dec.TransformFinalBlock(CIPHERS[t],0,16);
                    bool ok=true;
                    for(int j=0;j<8;j++) if(pt[j]!=WANT[t][j]) { ok=false; break; }
                    if(ok){ matched[t]=true; n++; }
                }
            }
        }
        return n;
    }

    public static List<string> RunAll(int[] pids, byte[][] frags, byte[][] ciphers, byte[] iv,
                                      string[] known, out long mb) {
        FRAGS=frags; CIPHERS=ciphers; IV=iv;
        SUF=new byte[]{1,0,0,0,0x73,0x41,0x6c,0x54};
        WANT=new byte[frags.Length][];
        for(int t=0;t<frags.Length;t++){
            WANT[t]=new byte[8];
            for(int j=0;j<8;j++) WANT[t][j]=(byte)(frags[t][j]^iv[j]);
        }
        matched=new bool[frags.Length];
        AES=Aes.Create(); AES.Mode=CipherMode.ECB; AES.Padding=PaddingMode.None;
        var found=new List<string>(); mb=0; Limited=false;
        PreMatched = MarkKnown(known);
        var sw = System.Diagnostics.Stopwatch.StartNew();
        foreach(int pid in pids) {
            if (AllDone() || Limited) break;
            IntPtr h=OpenProcess(0x0010|0x0400, false, pid);
            if (h==IntPtr.Zero) continue;
            IntPtr addr=IntPtr.Zero; MBI mbi; uint sz=(uint)Marshal.SizeOf(typeof(MBI));
            long scanned=0;
            while (VirtualQueryEx(h, addr, out mbi, sz)!=0) {
                if (sw.Elapsed.TotalSeconds > MaxSeconds) { Limited=true; break; }
                ulong baseA=(ulong)mbi.BaseAddress.ToInt64();
                ulong size=(ulong)mbi.RegionSize.ToInt64();
                uint prot = mbi.Protect & 0xFF;
                bool typeOk = WIDE || mbi.Type==MEM_PRIVATE;
                bool protOk = (prot==PAGE_RW || prot==PAGE_WC)
                              || (WIDE && (prot==0x02 || prot==0x40 || prot==0x80));
                bool ok = mbi.State==MEM_COMMIT && typeOk && protOk
                          && (mbi.Protect & (PAGE_GUARD|PAGE_NOACCESS))==0;
                if (ok && size>0 && size<0x40000000UL) {
                    try {
                        byte[] buf=new byte[size]; IntPtr rd;
                        if (ReadProcessMemory(h, mbi.BaseAddress, buf, (IntPtr)size, out rd)) {
                            int n=(int)rd; scanned+=n;
                            for (int o=0; o+16<=n; o+=8) {
                                if (!Candidate(buf,o)) continue;
                                foreach (string hit in Verify(buf,o)) found.Add(hit);
                                if (AllDone()) break;
                            }
                        }
                    } catch {}
                }
                ulong next=baseA+size; if(next<=baseA) break;
                addr=(IntPtr)(long)next; if(next>0x7FFFFFFFFFFFUL) break;
                if (AllDone() || sw.Elapsed.TotalSeconds > MaxSeconds) { if(!AllDone()) Limited=true; break; }
            }
            CloseHandle(h); mb+=scanned/1048576;
        }
        ElapsedSeconds = sw.Elapsed.TotalSeconds;
        return found;
    }
}
'@
Add-Type -TypeDefinition $cs -Language CSharp
[MultiKeyScan]::EnableDebug()

# 1) 收集每个账号 message.db 的库头校验靶
$wxRoot = Join-Path $env:USERPROFILE 'Documents\WXWork'
$targets = @()
foreach ($f in (Get-ChildItem $wxRoot -Filter message.db -Recurse -Depth 3 -ErrorAction SilentlyContinue)) {
    if ($f.FullName -match '\\Backup\\') { continue }
    try {
        $fsr = [System.IO.File]::Open($f.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $h = New-Object byte[] 96
        $fsr.Read($h, 0, 96) | Out-Null
        $fsr.Close()
        $rel = $f.FullName.Substring($wxRoot.Length).TrimStart('\')
        $account = ($rel -split '\\')[0]
        $targets += [pscustomobject]@{
            Account = $account
            Path    = $f.FullName
            Frag    = [byte[]]($h[16..23])
            Cipher  = [byte[]]($h[8..15] + $h[24..31])
        }
    } catch {}
}
if ($targets.Count -eq 0) { Write-Output 'NO_MESSAGE_DB'; exit 1 }

$frags   = New-Object 'byte[][]' $targets.Count
$ciphers = New-Object 'byte[][]' $targets.Count
for ($i = 0; $i -lt $targets.Count; $i++) {
    $frags[$i]   = [byte[]]$targets[$i].Frag
    $ciphers[$i] = [byte[]]$targets[$i].Cipher
}
$IV      = [byte[]](0x20,0xd7,0x42,0x0f,0x9c,0x37,0xa3,0x5d,0xca,0x6f,0xe9,0x2a,0x1c,0x69,0x99,0xa9)

Write-Output "目标账号数: $($targets.Count)"
$outFile = Join-Path $here 'decrypt\windows\wxwork_keys.json'
$knownKeys = @()
if (Test-Path $outFile) {
    try {
        $oldJson = Get-Content $outFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($prop in $oldJson.PSObject.Properties) {
            if ($prop.Name -and $prop.Value) { $knownKeys += [string]$prop.Value }
        }
    } catch {}
}
$pids = @(Get-Process -Name WXWork -ErrorAction SilentlyContinue | Sort-Object WorkingSet64 -Descending | ForEach-Object { $_.Id })
if ($pids.Count -eq 0) { Write-Output 'NO_WXWORK_PROCESS'; exit 1 }
Write-Output "WXWork 进程: $($pids -join ',')"

$scanSeconds = if ($env:WECOM_SCAN_SECONDS) { [int]$env:WECOM_SCAN_SECONDS } else { 120 }
[MultiKeyScan]::SetWide([bool]$Wide)
[MultiKeyScan]::SetMaxSeconds($scanSeconds)
Write-Output "扫描模式: $(if ([bool]$Wide) { 'WIDE' } else { 'NORMAL' }) / 时间上限: ${scanSeconds}s"
$mb = 0
$hits = [MultiKeyScan]::RunAll([int[]]$pids, $frags, $ciphers, $IV, [string[]]$knownKeys, [ref]$mb)
Write-Output "已知 key 预匹配: $([MultiKeyScan]::PreMatched)/$($targets.Count)"
Write-Output "扫描内存: ${mb}MB / 耗时 $([math]::Round([MultiKeyScan]::ElapsedSeconds,1))s"
if ([MultiKeyScan]::Limited) { Write-Output 'SCAN_LIMITED=1（达到时间上限，已保留本次与历史找到的 key）' }

$map = [ordered]@{}
foreach ($hit in $hits) {
    $parts = $hit.Split('|')
    $hex = $parts[0]
    $idx = [int]$parts[1]
    $acc = $targets[$idx].Account
    $map[$acc] = $hex
    Write-Output "KEY=$hex ACCOUNT=$acc"
}

# 合并历史已抓到的 key，避免宽扫描时丢失之前的结果
if (Test-Path $outFile) {
    try {
        $old = Get-Content $outFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($prop in $old.PSObject.Properties) {
            if ($prop.Name -and $prop.Value -and -not $map.Contains($prop.Name)) { $map[$prop.Name] = $prop.Value }
        }
    } catch {}
}
$json = $map | ConvertTo-Json -Depth 4
[System.IO.File]::WriteAllText($outFile, $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Output "KEYS_FILE=$outFile"
Write-Output "MATCHED=$($map.Count)/$($targets.Count)"
