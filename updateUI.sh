#!/bin/bash
cd $(dirname $0)
rm -rf dist/*
cp original.xlsm dist/rick_toolbox.xlsm
zip dist/rick_toolbox.xlsm ladyrickUI.xml
zip dist/rick_toolbox.xlsm _rels/.rels
echo "Finish. Please add callback."
