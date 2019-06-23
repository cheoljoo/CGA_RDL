all:
	mkdir -p ./OUTPUT
	mkdir -p ./OUTPUT/tmp
	mkdir -p ./OUTPUT/stc
	perl 1_excel.pl --input=1_example.xlsx
	perl 2_replace.pl default.GV 2_example.cpp.stc
	perl 3_recover.pl --template=./OUTPUT/stc/src/2_example.cpp --working=./3_working.cpp.data
test: test.pl test.xlsx
	perl test.pl --input=test.xlsx
	perl 2_replace.pl default.GV test.cpp.stc
test_recover: 3_recover.sh
	sh 3_recover.sh OUTPUT/stc test_recover/working test_recover/merged
3:
	perl 3_recover.pl --template=./OUTPUT/stc/src/2_example.cpp --working=./3_working.cpp.data --merge=./c/d/a.cpp
clean:
	rm -rf *.log
	rm -rf OUTPUT
	rm -rf *.GV
	rm -rf 3_merge.cpp
