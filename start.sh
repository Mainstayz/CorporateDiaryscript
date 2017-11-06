#!/bin/bash

basepath=$(cd `dirname $0`; pwd)
rm -rf $basepath/result.xml
# Kudou 的SVN 路径 ，保证 svn log   
# 100 拉100条数据
svn log /Users/pillar/Desktop/work/Kudou -l 100 --xml >> $basepath/result.xml
# xyh svn 的名字
# 5 天
python $basepath/script.py xyh 5 $basepath/result.xml