#!/bin/bash

start=$(date +%s.%N)
# ls >& /dev/null    # hey, be quite, do not output to console....


start_s=$(echo $start | cut -d '.' -f 1)
start_ns=$(echo $start | cut -d '.' -f 2)

for (( i=0; i<=60; i++ ))
do
    ./demopso
done
end=$(date +%s.%N)
end_s=$(echo $end | cut -d '.' -f 1)
end_ns=$(echo $end | cut -d '.' -f 2)


time=$(( ( 10#$end_s - 10#$start_s ) * 1000 + ( 10#$end_ns / 1000000 - 10#$start_ns / 1000000 ) ))


echo "$time ms"