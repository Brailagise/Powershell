    if (-not ([System.Management.Automation.PSTypeName]"TrustEverything").Type)
    {
        Add-Type -TypeDefinition  @"
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
public static class TrustEverything
{
    private static bool ValidationCallback(object sender, X509Certificate certificate, X509Chain chain,
        SslPolicyErrors sslPolicyErrors) { return true; }
    public static void SetCallback() { System.Net.ServicePointManager.ServerCertificateValidationCallback = ValidationCallback; }
    public static void UnsetCallback() { System.Net.ServicePointManager.ServerCertificateValidationCallback = null; }
}
"@
    }
    [TrustEverything]::SetCallback()


$WirelessAdapter = Get-NetAdapter -Name *Wi-Fi*
$ExtremeMAC = $WirelessAdapter.MacAddress
$ExtremeLogon = Invoke-WebRequest -Uri "https://*snip*/mapi.fcgi" -Method "POST" -Headers @{"Cookie"="sid="; "Origin"="https://*snip*"; "Accept-Encoding"="gzip, deflate, br"; "Accept-Language"="en-US,en;q=0.9"; "X_SESSION_ID"="SESSION_NOT_SET"; "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36"; "Accept"="*/*"; "Referer"="https://*snipped*/MainApp.swf"; "X-Requested-With"="ShockwaveFlash/32.0.0.171"} -ContentType "text/xml" -Body ([System.Text.Encoding]::UTF8.GetBytes("<methodCall>$([char]10)  <methodName>login</methodName>$([char]10)  <params>$([char]10)    <param>$([char]10)      <value>$([char]10)        <string>*snipped*</string>$([char]10)      </value>$([char]10)    </param>$([char]10)    <param>$([char]10)      <value>$([char]10)        <string>*snipped*</string>$([char]10)      </value>$([char]10)    </param>$([char]10)  </params>$([char]10)</methodCall>")) -SessionVariable ExtremeSession
Sleep -Seconds 5
$ExtremeSessionID = (($ExtremeLogon.Content -split ";")[16]).TrimEnd('t','l','&')
$ExtremeCookie = "sid=" + ($ExtremeSession.Cookies.GetCookies("https://*snipped*")).Value
$ExtremeAddition = Invoke-WebRequest -Uri "https://*snipped*/mapi.fcgi" -Method "POST" -Headers @{"Cookie"="$ExtremeCookie"; "Origin"="https://*snipped*"; "Accept-Encoding"="gzip, deflate, br"; "Accept-Language"="en-US,en;q=0.9"; "X_SESSION_ID"="$ExtremeSessionID"; "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36"; "Accept"="*/*"; "Referer"="https://*snipped*/MainApp.swf"; "X-Requested-With"="ShockwaveFlash/32.0.0.171"} -ContentType "text/xml" -Body ([System.Text.Encoding]::UTF8.GetBytes("<methodCall>$([char]10)  <methodName>sendRequest</methodName>$([char]10)  <params>$([char]10)    <param>$([char]10)      <value>$([char]10)        <string>&lt;rpc message-id=`"24`"&gt;&lt;edit-config&gt;&lt;target&gt;&lt;running/&gt;&lt;/target&gt;&lt;config&gt;&lt;wing-config&gt;&lt;radius_user_pool&gt;&lt;users operation=`"create`"&gt;&lt;password&gt;$ExtremeMAC&lt;/password&gt;&lt;userid&gt;$ExtremeMAC&lt;/userid&gt;&lt;/users&gt;&lt;name&gt;*snipped*&lt;/name&gt;&lt;/radius_user_pool&gt;&lt;/wing-config&gt;&lt;/config&gt;&lt;/edit-config&gt;&lt;/rpc&gt;</string>$([char]10)      </value>$([char]10)    </param>$([char]10)  </params>$([char]10)</methodCall>"))
$ExtremeCommitandsave = Invoke-WebRequest -Uri "https://*snipped*/mapi.fcgi" -Method "POST" -Headers @{"Cookie"="$ExtremeCookie"; "Origin"="https://*snipped*"; "Accept-Encoding"="gzip, deflate, br"; "Accept-Language"="en-US,en;q=0.9"; "X_SESSION_ID"="$ExtremeSessionID"; "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36"; "Accept"="*/*"; "Referer"="https://*snipped*/MainApp.swf"; "X-Requested-With"="ShockwaveFlash/32.0.0.171"} -ContentType "text/xml" -Body ([System.Text.Encoding]::UTF8.GetBytes("<methodCall>$([char]10)  <methodName>sendRequest</methodName>$([char]10)  <params>$([char]10)    <param>$([char]10)      <value>$([char]10)        <string>&lt;rpc message-id=`"20`"&gt;&lt;commit-config&gt;&lt;target&gt;&lt;running/&gt;&lt;/target&gt;&lt;/commit-config&gt;&lt;copy-config&gt;&lt;source&gt;&lt;running/&gt;&lt;/source&gt;&lt;target&gt;&lt;memory/&gt;&lt;/target&gt;&lt;/copy-config&gt;&lt;/rpc&gt;</string>$([char]10)      </value>$([char]10)    </param>$([char]10)  </params>$([char]10)</methodCall>"))
$ExtremeLogoff = Invoke-WebRequest -Uri "https://*snipped*/mapi.fcgi" -Method "POST" -Headers @{"Cookie"="$ExtremeCookie"; "Origin"="https://*snipped*"; "Accept-Encoding"="gzip, deflate, br"; "Accept-Language"="en-US,en;q=0.9"; "X_SESSION_ID"="$ExtremeSessionID"; "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36"; "Accept"="*/*"; "Referer"="https://*snipped*/MainApp.swf"; "X-Requested-With"="ShockwaveFlash/32.0.0.171"} -ContentType "text/xml" -Body ([System.Text.Encoding]::UTF8.GetBytes("<methodCall>$([char]10)  <methodName>logout</methodName>$([char]10)</methodCall>"))

