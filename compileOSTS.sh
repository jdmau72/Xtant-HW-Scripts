#!/bin/bash
# Script run to create .osts file versions of each of the .ts scripts
# It's easier to work in VSCode than Excel's built-in scripting environment

src_dir=./ts
for src_file_path in "$src_dir"/*
do 
    src_file=$(basename "$src_file_path" .ts)
    echo "Copying $src_file ..."
    cp $src_file_path "./osts/$src_file.osts"
done
