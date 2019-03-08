#!/bin/bash
#

git pull

python3 gcal-sync.py >> sync.log

DUPS=`grep 'total events: 5' sync.log |wc -l`
echo Dups: $DUPS

