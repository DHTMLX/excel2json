#!/bin/bash
set -e

server=$1
folder=$2
version=$3

if [ -z "$server" ] || [ -z "$folder" ] || [ -z "$version" ]; then
    echo "Usage: $0 <server> <folder> <version>"
    echo "Example: $0 user@example.com /var/www/apps 1.3.0"
    exit 1
fi

# Extract subversion by dropping the last digit (e.g., 1.3.0 -> 1.3)
subversion=$(echo "$version" | sed 's/\.[0-9]*$//')

echo "Deploying version: $version"
echo "Creating subversion link: $subversion"

# Create version directory and upload files
ssh $server "mkdir -p ${folder}/${version}"
scp -r dist/worker.js "$server:${folder}/${version}/worker.js"
scp -r dist/module.js "$server:${folder}/${version}/module.js"
scp -r pkg/excel2json_wasm_bg.wasm "$server:${folder}/${version}/excel2json_wasm_bg.wasm"

# Remove existing subversion link and create new one
ssh $server "cd \"${folder}\" && rm -f \"./${subversion}\" && ln -s \"${version}\" \"${subversion}\""

echo "Deployment complete!"
