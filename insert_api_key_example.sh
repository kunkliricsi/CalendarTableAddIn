#!/bin/bash

find ./CalendarTableAddIn -type f \( ! -iname "*.sh" ! -iname "*.md" ! -iname "*.png" \) -exec sed -i -e "s/YOUR_API_KEY_GOES_HERE/$1/g" {} \;