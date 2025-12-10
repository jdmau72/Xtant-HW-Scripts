#!/bin/bash
# alternate script in case I want to quickly generate the .ts scripts from a current .osts

src_dir=./osts
src_file=$(basename $1 .osts)
cp "$src_dir"/$1 "./ts/$src_file.ts"

