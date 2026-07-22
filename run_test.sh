#!/usr/bin/env bash
# Runs digit_testplan.jmx headlessly, stops automatically at 2% error rate
# (handled inside the .jmx by the "AutoStop on 2% Error Rate" JSR223 Listener),
# then converts the raw results into an Excel aggregate report.
#
# Usage: ./run_test.sh

set -euo pipefail

JMX_FILE="digit_testplan.jmx"
TIMESTAMP=$(date +%Y%m%d_%H%M%S)
RESULTS_JTL="results_${TIMESTAMP}.jtl"
JMETER_LOG="jmeter_${TIMESTAMP}.log"
AGGREGATE_XLSX="aggregate_report_${TIMESTAMP}.xlsx"

echo "Starting JMeter run: ${JMX_FILE}"
jmeter -n -t "${JMX_FILE}" \
  -l "${RESULTS_JTL}" \
  -j "${JMETER_LOG}"

echo "JMeter run finished. Building Excel aggregate report..."
python3 generate_aggregate_excel.py "${RESULTS_JTL}" "${AGGREGATE_XLSX}"

echo "Done."
echo "  Raw results : ${RESULTS_JTL}"
echo "  JMeter log  : ${JMETER_LOG}"
echo "  Aggregate   : ${AGGREGATE_XLSX}"
