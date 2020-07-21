#!/usr/bin/env bash
docker rmi -f petrjahoda/likovspecialservice:latest
docker build -t petrjahoda/likovspecialservice:latest .
docker push petrjahoda/likovspecialservice:latest