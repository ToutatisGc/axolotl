#!/bin/bash

if ! command -v docsify > /dev/null 2>&1; then
  echo "Error: [docsify-cli]未安装,请使用[npm install docsify-cli -g]安装."
  exit 1
fi

docsify serve ../..