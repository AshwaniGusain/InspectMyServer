#!/bin/bash

# Command to fetch processor information
cpu_info=$(lscpu)

# Extracting relevant information from the output
name=$(grep "Model name:" <<< "$cpu_info" | awk -F ': ' '{print $2}')
manufacturer=$(grep "Vendor ID:" <<< "$cpu_info" | awk -F ': ' '{print $2}')
l2_cache=$(grep "L2 cache:" <<< "$cpu_info" | awk -F ': ' '{print $2}')
ext_clock=$(grep "CPU MHz:" <<< "$cpu_info" | awk -F ': ' '{print $2}')
load=$(uptime | awk -F 'load average:' '{print $2}' | awk '{print $3}' | sed 's/,//')
num_cores=$(nproc)
num_logical_processors=$(nproc --all)

# Command to fetch operating system information
os_version=$(lsb_release -d | awk -F ':\t' '{print $2}')
logged_in_user=$(whoami)

# Command to fetch open network connections
open_connections=$(ss -tuln state all)

# Command to fetch system information
computer_name=$(hostname)
model=$(dmidecode -t system | grep "Product Name" | awk -F ': ' '{print $2}')
user_name=$(whoami)
ip_address=$(hostname -I | awk '{print $1}')

# Command to fetch running processes information
running_processes=$(ps -eo pid,comm,cmd)

# Command to fetch installed applications information
installed_apps=$(dpkg -l | grep '^ii' | awk '{print $2,$3}')

# HTML content for the report
html_report=$(cat <<EOF
<table align="center" style="width: 900px;WORD-BREAK:BREAK-ALL">
<tr>
  <td align="center" bgcolor="#AAAAAA" colspan="5">
    <b>Processor Information</b><font size="1"> - <a href="http://www.lookinmypc.com/help.htm#50" target="_blank">What's This?</a></font>
  </td>
</tr>
<tr>
  <td align="center" bgcolor="#C6D8EC">Name</td>
  <td align="center" bgcolor="#C6D8EC">Manufacturer</td>
  <td align="center" bgcolor="#C6D8EC">L2 Cache</td>
  <td align="center" bgcolor="#C6D8EC">Ext Clock</td>
  <td align="center" bgcolor="#C6D8EC">Load</td>
</tr>
<tr>
  <td align="center">$name</td>
  <td align="center">$manufacturer</td>
  <td align="center">$l2_cache</td>
  <td align="center">$ext_clock</td>
  <td align="center">$load</td>
</tr>
</table>

<table align="center" style="width: 900px;WORD-BREAK:BREAK-ALL">
<tr>
  <td align="center" bgcolor="#AAAAAA" colspan="3">
    <b>CPU Information</b>
  </td>
</tr>
<tr>
  <td align="center" bgcolor="#C6D8EC">Name</td>
  <td align="center" bgcolor="#C6D8EC">Number of Cores</td>
  <td align="center" bgcolor="#C6D8EC">Number of Logical Processors</td>
</tr>
<tr>
  <td align="center">$name</td>
  <td align="center">$num_cores</td>
  <td align="center">$num_logical_processors</td>
</tr>
</table>

<table align="center" style="width: 900px;WORD-BREAK:BREAK-ALL">
<tr>
  <td align="center" bgcolor="#AAAAAA" colspan="9">
    <b>Open Network Connections</b>
  </td>
</tr>
<tr>
  <td align="center" bgcolor="#C6D8EC">PID</td>
  <td align="center" bgcolor="#C6D8EC">Process Name</td>
  <td align="center" bgcolor="#C6D8EC">Protocol</td>
  <td align="center" bgcolor="#C6D8EC">Local Address</td>
  <td align="center" bgcolor="#C6D8EC">Local Port</td>
  <td align="center" bgcolor="#C6D8EC">Remote Address</td>
  <td align="center" bgcolor="#C6D8EC">Remote Port</td>
  <td align="center" bgcolor="#C6D8EC">State</td>
</tr>
EOF
)

# Add open network connections to the HTML report
while read -r line; do
    html_report+="<tr><td align='center'>$(awk '{print $2}' <<< "$line")</td>"
    html_report+="<td align='center'>$(awk '{print $6}' <<< "$line")</td>"
    html_report+="<td align='center'>$(awk '{print $1}' <<< "$line")</td>"
    html_report+="<td align='center'>$(awk '{print $4}' <<< "$line")</td>"
    html_report+="<td align='center'>$(awk -F':' '{print $2}' <<< "$(awk '{print $4}' <<< "$line")")</td>"
    html_report+="<td align='center'>$(awk '{print $5}' <<< "$line")</td>"
    html_report+="<td align='center'>$(awk -F':' '{print $2}' <<< "$(awk '{print $5}' <<< "$line")")</td>"
    html_report+="<td align='center'>$(awk '{print $6}' <<< "$line")</td></tr>"
done <<< "$open_connections"

html_report+="</table>"

# Add operating information to the HTML report
html_report+=$(cat <<EOF
<table align="center" style="width: 900px;WORD-BREAK:BREAK-ALL">
<tr>
  <td align="center" bgcolor="#AAAAAA" colspan="2">
    <b>Operating Information</b>
  </td>
</tr>
<tr>
  <td align="center" bgcolor="#C6D8EC">Operating System Version</td>
  <td align="center" bgcolor="#C6D8EC">Logged In User</td>
</tr>
<tr>
  <td align="center">$os_version</td>
  <td align="center">$logged_in_user</td>
</tr>
</table>
EOF
)

# Add system information to the HTML report
html_report+=$(cat <<EOF
<table align="center" style="width: 900px;WORD-BREAK:BREAK-ALL">
<tr>
  <td align="center" bgcolor="#AAAAAA" colspan="2">
    <b>System Information</b>
  </td>
</tr>
<tr>
  <td align="center" bgcolor="#C6D8EC">Computer Name</td>
  <td align="center" bgcolor="#C6D8EC">Manufacturer</td>
  <td align="center" bgcolor="#C6D8EC">Model</td>
  <td align="center" bgcolor="#C6D8EC">User Name</td>
  <td align="center" bgcolor="#C6D8EC">IP Address</td>
</tr>
<tr>
  <td align="center">$computer_name</td>
  <td align="center">$manufacturer</td>
  <td align="center">$model</td>
  <td align="center">$user_name</td>
  <td align="center">$ip_address</td>
</tr>
</table>
EOF
)

# Add running processes information to the HTML report
html_report+=$(cat <<EOF
<table align="center" style="width: 900px;WORD-BREAK:BREAK-ALL">
<tr>
  <td align="center" bgcolor="#AAAAAA" colspan="3">
    <b>Running Processes</b>
  </td>
</tr>
<tr>
  <td align="center" bgcolor="#C6D8EC">PID</td>
  <td align="center" bgcolor="#C6D8EC">Name</td>
  <td align="center" bgcolor="#C6D8EC">Path</td>
</tr>
EOF
)

while read -r line; do
    pid=$(awk '{print $1}' <<< "$line")
    name=$(awk '{print $2}' <<< "$line")
    path=$(awk '{$1="";$2="";print}' <<< "$line")
    html_report+="<tr><td align='center'>$pid</td><td align='center'>$name</td><td>$path</td></tr>"
done <<< "$running_processes"

html_report+="</table>"

# Add installed applications information to the HTML report
html_report+=$(cat <<EOF
<table align="center" style="width: 900px;WORD-BREAK:BREAK-ALL">
<tr>
  <td align="center" bgcolor="#AAAAAA" colspan="2">
    <b>Installed Applications</b>
  </td>
</tr>
<tr>
  <td align="center" bgcolor="#C6D8EC">Name</td>
  <td align="center" bgcolor="#C6D8EC">Version</td>
</tr>
EOF
)

while read -r line; do
    name=$(awk '{print $1}' <<< "$line")
    version=$(awk '{print $2}' <<< "$line")
    html_report+="<tr><td align='center'>$name</td><td align='center'>$version</td></tr>"
done <<< "$installed_apps"

html_report+="</table>"

# Output the HTML content to a file (e.g., report.html)
output_file="report.html"
echo "$html_report" > "$output_file"

echo "HTML report generated successfully. Please check $output_file."


sender="ashwanigusain@live.com"
# Recipient's email address
recipient="ashwanigusain@live.com"

# Subject of the email
subject="Automated Email with HTML Attachment"

# HTML content file
html_file="$output_file"

# Create the email body
email_body="Please find attached an HTML file."

# Create a temporary text file for the email body
email_body_file=$(mktemp)
echo -e "$email_body" > "$email_body_file"

# Send the email
sendmail -t <<END
From:@sender
To: $recipient
Subject: $subject
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="XYZ"

--XYZ
Content-Type: text/plain; charset=us-ascii
Content-Disposition: inline
Content-Transfer-Encoding: quoted-printable

$(cat "$email_body_file")

--XYZ
Content-Type: text/html; charset=us-ascii
Content-Disposition: attachment; filename="email_content.html"
Content-Transfer-Encoding: base64

$(base64 "$html_file")

--XYZ--
END

echo "Heloo I'm here"
# Clean up temporary files
rm "$email_body_file"

if [ $? -eq 0 ]; then
    echo "Email sent successfully"
else
    echo "Email sending failed"
fi

















