#!/bin/bash

for py_file in $(find $Codes -name '*.py')
do
    python $py_file
done
