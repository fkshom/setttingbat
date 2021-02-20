# ref
# https://github.com/Disassembler0/Win10-Initial-Setup-Script/issues/229

function Refresh-Explorer {
    if ( -not ([System.Management.Automation.PSTypeName]'WindowsDesktopTools.Explorer').Type) {
        $typeParams = @{
            Namespace = 'WindowsDesktopTools'
            Name = 'Explorer'
            Language = 'CSharp'
            MemberDefinition = @'
                private static readonly IntPtr HWND_BROADCAST = new IntPtr(0xffff);
                private const int WM_SETTINGCHANGE = 0x1a;
                private const int SMTO_ABORTIFHUNG = 0x0002;
                
                [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
                static extern bool SendNotifyMessage(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam);

                [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
                private static extern IntPtr SendMessageTimeout(IntPtr hWnd, int Msg, IntPtr wParam, string lParam, int fuFlags, int uTimeout, IntPtr lpdwResult);

                [DllImport("shell32.dll", CharSet = CharSet.Auto, SetLastError = false)]
                private static extern int SHChangeNotify(int eventId, int flags, IntPtr item1, IntPtr item2);

                public static void Refresh()
                {
                    // Update desktop icons
                    SHChangeNotify(0x8000000, 0x1000, IntPtr.Zero, IntPtr.Zero);
                    // Update environment variables
                    SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, IntPtr.Zero, null, SMTO_ABORTIFHUNG, 100, IntPtr.Zero);
                    // Update taskbar
                    SendNotifyMessage(HWND_BROADCAST, WM_SETTINGCHANGE, IntPtr.Zero, "TraySettings");
                    
                }
'@
        }
        Add-Type @typeParams -IgnoreWarnings -ErrorAction Stop
    }
    Write-Verbose 'Refreshing Shell environment ...'
    [WindowsDesktopTools.Explorer]::Refresh()
}

Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced -Name TaskbarSmallIcons -Value 1
Refresh-Explorer
