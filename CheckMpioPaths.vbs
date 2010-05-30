Sub NagiosExit(warning_messages, num_warning_messages, _
  critical_messages, num_critical_messages)
    Dim status_str
    status_str = ""
    
    If num_critical_messages > 0 Then
        For i = 0 To num_critical_messages
            status_str = status_str + critical_messages(i)
            
            If i < num_critical_messages - 1 Then
                status_str = status_str + "; "
            End If
        Next
        
        WScript.Echo "MPIO CRITICAL: " + status_str
        WScript.Quit 2
    End If
    
    If num_warning_messages > 0 Then
        For i = 0 To num_warning_messages
            status_str = status_str + critical_messages(i)
            
            If i < num_warning_messages - 1 Then
                status_str = status_str + "; "
            End If
        Next
        
        WScript.Echo "MPIO WARNING: " + status_str
        WScript.Quit 1    
    End If
    
    WScript.Echo "MPIO OK: All disks are within path thresholds."
    WScript.Quit 0
End Sub

Dim computer
computer = "."

Dim min_allowed_paths
min_allowed_paths = 4

Dim linked_opt
linked_opt = ""

' Parse command line options
For Each opt In WScript.Arguments
    If linked_opt = "/paths" Then
        min_allowed_paths = CInt(opt)
        linked_opt = ""
    End If
    
    If opt = "/paths" Then
        linked_opt = "/paths"
    End If
Next

Dim wmi
Set wmi = GetObject("winmgmts://" + computer + "/root/WMI")

Dim warning_messages(100)
Dim num_warning_messages
num_warning_messages = 0

Dim critical_messages(100)
Dim num_critical_messages
num_critical_messages = 0

Dim mpio_disks
Set mpio_disks = wmi.ExecQuery("SELECT * FROM MPIO_DISK_INFO")
For Each disk In mpio_disks
    Dim mpio_drives
    mpio_drives = disk.DriveInfo
    
    For Each drive In mpio_drives
        Dim name
        name = drive.Name
        
        Dim paths
        paths = drive.NumberPaths
        
        If drive.NumberPaths < min_allowed_paths Then
            critical_messages(num_critical_messages) = _
              "Drive """ + name + """ has " + CStr(paths) + " paths"
            num_critical_messages = num_critical_messages + 1
        End If
    Next
Next

NagiosExit warning_messages, num_warning_messages, critical_messages, _
  num_critical_messages
