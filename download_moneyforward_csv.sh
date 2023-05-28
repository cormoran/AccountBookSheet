#!/bin/bash

for year in $(seq 2020 2021); do
    for month in $(seq -f '%02g' 12); do
        open "https://moneyforward.com/cf/csv?from=$year%2F$month%2F01&month=$month&year=$year"
    done
done
