# Pass-Patrol

A tool to search for secrets in directories.

Handle pdf and some office documents

/!\ relatively slow compared to a grep for example


egrep -ari --exclude-from ext.grep  "passw|credential|creds" ~/Downloads


egrep -ari --exclude-from=ext.grep "passw|credential|creds" ~/Downloads | awk -F: '{m=match($0, /passw|credential|creds/);print "\033[1;35m"$1"\033[0m:"substr($0,(m>40?m-40:0),40)"\033[1;31m"substr($0,m,RLENGTH)"\033[0m"substr($0,m+RLENGTH,40)}'
