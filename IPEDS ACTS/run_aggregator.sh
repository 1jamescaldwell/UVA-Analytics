#!/bin/bash
EXEC="path/to/aggregator_executable" # Replace with the actual path to the executable

BASE_DIR="path/to/template_files" # Replace with the actual base directory containing the template files

run_case () {
  echo "Running year $2"

  "$EXEC" <<EOF
$1
$2
$3
$4
$5
$6
EOF

  STATUS=$?
  if [ $STATUS -ne 0 ]; then
    echo "ERROR on year $2 — stopping script."
    exit 1
  fi
}

run_case "234076" \
"2019-20" \
"$BASE_DIR/1920/ACTS_Template_AY2019-20.xlsx" \
"$BASE_DIR/1920" \
"yes" \
"yes"

run_case "234076" \
"2020-21" \
"$BASE_DIR/2021/ACTS_Template_AY2020-21.xlsx" \
"$BASE_DIR//2021" \
"yes" \
"yes"

run_case "234076" \
"2021-22" \
"$BASE_DIR/2122/ACTS_Template_AY2021-22.xlsx" \
"$BASE_DIR/2122" \
"yes" \
"yes"

run_case "234076" \
"2022-23" \
"$BASE_DIR/2223/ACTS_Template_AY2022-23.xlsx" \
"$BASE_DIR/2223" \
"yes" \
"yes"

run_case "234076" \
"2023-24" \
"$BASE_DIR/2324/ACTS_Template_AY2023-24.xlsx" \
"$BASE_DIR/2324" \
"yes" \
"yes"

run_case "234076" \
"2024-25" \
"$BASE_DIR/2425/ACTS_Template_AY2024-25.xlsx" \
"$BASE_DIR/2425" \
"yes" \
"yes"

run_case "234076" \
"2025-26" \
"$BASE_DIR/2526/ACTS_Template_AY2025-26.xlsx" \
"$BASE_DIR/2526" \
"yes" \
"yes"