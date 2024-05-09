#! /bin/bash

# variables
timestamp=$(date +"%Y-%m-%d_%H-%M-%S")
current_dir=$(pwd)
path=$1
cracked_output="CRACKED_$path_$timestamp.xlsm"

# Set an environment variable with the current working directory
export MY_SCRIPT_DIR="$current_dir"

# Define ANSI escape codes for green color
GREEN='\033[0;32m'
YELLOW='\033[0;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

# Check if a filename was provided as a command-line argument
if [ $# -eq 0 ]; then
    echo "Usage: $0 <filename>"
    exit 1
fi

# Check if the file exists
if [ ! -f "$path" ]; then
    echo -e "${RED}[-] File '$path' not found.${NC}"
    exit 1
fi
echo -e "${GREEN}[+] Processing file: $path ${NC}"

# Check for required hex editor
if command -v xxd &> /dev/null; then
    echo -e "${GREEN}[+] xxd is installed."${NC}
else
    echo -e "${RED}[-] xxd is not installed. Please install it with 'sudo apt install xxd'"${NC}
    exit 1
fi

# Copy to staging directory
echo -e "${GREEN}[+] copying $path to staging directory..."${NC}
mkdir $MY_SCRIPT_DIR/staging
cp $path $MY_SCRIPT_DIR/staging/tobe-cracked.zip

# extract VBA file contents
echo -e "${GREEN}[+] extracting .xlsm file..."${NC}
unzip $MY_SCRIPT_DIR/staging/tobe-cracked.zip -d $MY_SCRIPT_DIR/staging > /dev/null
rm $MY_SCRIPT_DIR/staging/tobe-cracked.zip

# swap DPB hex bit to DPX
echo -e "${GREEN}[+] breaking password protection mechanism...${NC}"
xxd -p $MY_SCRIPT_DIR/staging/xl/vbaProject.bin > $MY_SCRIPT_DIR/temp.hex
sed -i 's/445042/445058/g' $MY_SCRIPT_DIR/temp.hex
xxd -r -p $MY_SCRIPT_DIR/temp.hex > $MY_SCRIPT_DIR/staging/xl/vbaProject.bin
rm $MY_SCRIPT_DIR/temp.hex

# repackage VBA contents
echo -e "${GREEN}[+] packaging file back into .xlsm${NC}"
cd $MY_SCRIPT_DIR/staging
zip -r $MY_SCRIPT_DIR/cracked.zip ./* > /dev/null
cd $MY_SCRIPT_DIR
mv $MY_SCRIPT_DIR/cracked.zip "$MY_SCRIPT_DIR/$cracked_output"

# remove artifacts
echo -e "${GREEN}[+] removing staging directory${NC}"
rm -rf $MY_SCRIPT_DIR/staging

echo -e "${GREEN}[+] Done! cracked file is "$MY_SCRIPT_DIR/$cracked_output"${NC}"
