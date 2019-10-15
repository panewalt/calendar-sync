#!/bin/bash
#

git pull

python3 gcal-sync.py pete >> pete/sync.log

DUPS=`grep 'total events: 5' pete/sync.log |wc -l`
echo Dups: $DUPS

