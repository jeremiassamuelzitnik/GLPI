# GLPI

Exe creation command:

ps2exe .\GLPI-agent-online-installer.ps1 -x64 -iconFile .\Icon.ico -title "GLPI online installer" -company "GLPI" -version 1.1.2 -requireAdmin

$setupOptions = '/quiet RUNNOW=1 SERVER=http://YOUR_SERVER/marketplace/glpiinventory/ ADD_FIREWALL_EXCEPTION=1 ADDLOCAL=feat_AGENT,feat_DEPLOY EXECMODE=1'