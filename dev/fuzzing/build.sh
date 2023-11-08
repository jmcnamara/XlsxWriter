cd "$SRC"/XlsxWriter
pip3 install .

# Build fuzzers in $OUT
for fuzzer in $(find dev/fuzzing -name '*_fuzzer.py');do
  compile_python_fuzzer "$fuzzer"
done
zip -q $OUT/xlsx_fuzzer_seed_corpus.zip $SRC/corpus/*
