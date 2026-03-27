@{
    # ---------------------------------------------------------------------------
    # Sliding window size in minutes.
    # The script counts how many times a given Event ID / Source combination
    # occurs within the last CompareInterval minutes. If the count reaches
    # CompareCount, a Zabbix trap is sent.
    # ---------------------------------------------------------------------------
    CompareInterval = 10

    # ---------------------------------------------------------------------------
    # Minimum number of occurrences within CompareInterval minutes that triggers
    # a Zabbix trap. Increase this to reduce noise on busy servers.
    # ---------------------------------------------------------------------------
    CompareCount = 5

    # ---------------------------------------------------------------------------
    # Cooldown period in minutes. After a trap is sent for a given Event ID /
    # Source pair, no further traps will be sent for that pair until FloodFuse
    # minutes have elapsed. This prevents alert storms.
    # Default: 1440 (24 hours). Set to 0 to disable flood protection.
    # ---------------------------------------------------------------------------
    FloodFuse = 1440

    # ---------------------------------------------------------------------------
    # Full path to the zabbix_sender executable on this host.
    # zabbix_sender is included with the Zabbix agent package.
    # ---------------------------------------------------------------------------
    ZabbixSenderPath = "C:\Program Files\ZabbixAgent\zabbix_sender.exe"

    # ---------------------------------------------------------------------------
    # Full path to the Zabbix agent configuration file.
    # zabbix_sender reads it to determine the local host name and server address.
    # ---------------------------------------------------------------------------
    ZabbixConfigPath = "C:\Program Files\ZabbixAgent\zabbix_agentd.win.conf"

    # ---------------------------------------------------------------------------
    # Windows event log channels to monitor.
    # Add or remove channels as needed.
    # Examples: "Security", "Microsoft-Windows-Sysmon/Operational",
    #            "Microsoft-Windows-PowerShell/Operational"
    # ---------------------------------------------------------------------------
    EventLogs = @("System", "Application")
}
