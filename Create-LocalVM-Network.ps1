New-VMSwitch -SwitchName "Hyper-V Firewall Network" -SwitchType Internal
New-NetIPAddress -IPAddress 10.0.0.1 -PrefixLength 24 -InterfaceAlias "vEthernet (Hyper-V Firewall Network)"
New-NetNAT -Name "Hyper-V Firewall Network" -InternalIPInterfaceAddressPrefix 10.0.0.0/24
