#!/bin/bash
pkill -f soffice 2>/dev/null; sleep 1
rm -rf /tmp/lo_uno
exec /opt/libreoffice7.6/program/soffice --headless --norestore --nolockcheck   -env:UserInstallation=file:///tmp/lo_uno   --accept="socket,host=127.0.0.1,port=2002;urp;" >>/tmp/uno.log 2>&1
