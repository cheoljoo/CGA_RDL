#!/bin/perl
use 5.010;

use File::Path qw(make_path);
use Getopt::Long;
use File::Basename qw(dirname);
use Cwd  qw(abs_path);
use lib dirname(dirname abs_path $0) . '/perllib';



sub __SUB__ { return  (caller 2)[3] . "|" . (caller 2)[2] . "-" . (caller 1)[3] . "|" . (caller 1)[2] . "-" . (caller 0)[2] . ": " }

sub versionMismatch {
	print STDERR "Version Mismatch between excel file and this git repository\n";
	print STDERR "1.check Version of CGA_RDL in $cga_rdl_version_input_file file\n";
	print STDERR "Please update your CGA_RDL version and Excel version.\n";
    exit(4);
}


sub help 
{
	printf("Help :\n");
	printf("\t--inputstci=[stcI or stc file]\n");
	printf("\t\t  default input stc or stcI file name : $stcfilename\n");
	printf("\t--outputdb=[output file]\n");
	printf("\t\t  default output DB file name : $filename\n");
	printf("\t--cga_rdl_version_input=[version file name]\n");
	printf("\t\t  default input version file name : $cga_rdl_version_input_file\n");
	printf("\t\t  if null , we ignore the version between excel version input file and [VARIABLE]Excel_Version  in excel file ($cga_rdl_version_input_file)\n");
	printf("\t--debug");
	printf("\t\t  debug mode : ITERATE  IFEQUAL IFNOTEQUAL\n");
	printf("\t--debugdetail");
	printf("\t\t  --debug + ITERATE detail \n");
	printf("\t\t  debug mode : ITERATE  IFEQUAL IFNOTEQUAL  ITERATE detail \n");
	printf("\t--original");
	printf("\t\t  run NO performance mode\n");
	printf("\t--nolog");
	printf("\t\t  print log into /dev/null\n");
	printf("\t--help\n");
}

sub plus {
    my $num = shift @_;
    my $cnt = shift @_;
    $num += $cnt;
    return  $num;
}

sub minus {
    my $num = shift @_;
    my $cnt = shift @_;
    $num -= $cnt;
    return  $num;
}

sub ifthisthenthat {
    my $my_if = shift @_;
    my $my_this = shift @_;
    my $my_that = shift @_;
    #print STDERR "my_if : $my_if\n";
    #print STDERR "my_this : $my_this\n";
    if($my_if){ return $my_this; }
    else { return $my_that; }
}

sub substitute {
    my $my_s = shift @_;
    my $my_from = shift @_;
    my $my_to = shift @_;
    #print STDERR "my_from : $my_from\n";
    #print STDERR "my_to : $my_to\n";
    $my_s =~ s/$my_from/$my_to/g;
    return $my_s;
}

sub in {
    my $my_a = shift @_;
    my $my_b = shift @_;
    if($my_a = /$my_b/){
        return $my_b;
    } else {
        return "";
    }
}

sub countKey {
    my $my_s = shift @_;
    my $my_count = 0;
    foreach my $my_key (sort_keys($my_s)){
        #print STDERR "countKey : KEY : $my_key\n";
        $my_count++;
    }
    #print STDERR "countKey : $my_s : $my_count\n";
    return $my_count;
}

sub start_time_log {
	my $tmpLogInit = shift @_;
	($Second, $Minute, $Hour, $Day, $Month, $Year, $WeekDay, $DayOfYear, $IsDST) = localtime(time) ;
	$Month++;
	$Year += 1900;
	print_fp( "$tmpLogInit structg : CChange START ($stc_filename_input) - $Month-$Day-$Year : $Hour : $Minute : $Second\n",DBG,TIME_DBG);
	$tttime = $Hour * 3600 + $Minute * 60 + $Second;
	$cchange_start_time = $Hour * 3600 + $Minute * 60 + $Second;
}
sub mid_time_log {
	my $tmpLogInit = shift @_;
	($Second, $Minute, $Hour, $Day, $Month, $Year, $WeekDay, $DayOfYear, $IsDST) = localtime(time) ;
	$Month++;
	$Year += 1900;
	$ddtime = ($Hour * 3600 + $Minute * 60 + $Second) - $tttime;
	$tttime = $Hour * 3600 + $Minute * 60 + $Second;
	print_fp( "$tmpLogInit structg ($ddtime) : CChange ITERATE recursion END 1 ($stc_filename_input) - $Month-$Day-$Year : $Hour : $Minute : $Second\n",DBG,TIME_DBG);
}
sub end_time_log {
	my $tmpLogInit = shift @_;
	($Second, $Minute, $Hour, $Day, $Month, $Year, $WeekDay, $DayOfYear, $IsDST) = localtime(time) ;
	$Month++;
	$Year += 1900;
	$ddtime = ($Hour * 3600 + $Minute * 60 + $Second) - $cchange_start_time;
	print_fp( "$tmpLogInit structg CChange duration ($ddtime) : CChange END ($stc_filename_input) - $Month-$Day-$Year : $Hour : $Minute : $Second\n",DBG,TIME_DBG);
}

sub print_fp
{
	my $print_contents;
	my $to_file;

	$print_contents = shift @_;
	foreach $to_file (@_) {
		## STD라는 것은 DBG에 대해서 만들어둔 것으로 STD 나 DBG이나 같은 값을 의미한다.  
		print $to_file $print_contents;
	}
}

sub IChange
{
	my $iterate_key;
	my $iterate_value;
	my $iterate_hash_name;
	my $iterate_lines = "";
	my $in;
	my $in_start;
	my $in_end;
	my $stcI_filename_input ;
	my $stcI_filename_output ;
	my $stcI_output_dir;
	my $stcI_for;
	my $stcI_extension="";
	my $stcI_filepostfix="";
	my $stcI_fileprefix="";
	my $stcI_filebody="YES";
	my $stcI_file_lower = "NO";
    my $stcI_all_lower = "NO";
	my $stcI_lines = "";
	my @stcI_lines;
	my $stcfilename="";
	my $tttime;
	my $ddtime;

	$stcI_filename_input = shift @_;
	$stcI_output_dir = shift @_;

	print_fp("\n\n\nIchange $stcI_filename_input\n",STDOUT,DBG);
	open(STCI_INPUT , "<$stcI_filename_input");

($Second, $Minute, $Hour, $Day, $Month, $Year, $WeekDay, $DayOfYear, $IsDST) = localtime(time) ;
$Month++;
$Year += 1900;
print_fp( "structg : Ichange START ($stcI_filename_input) - $Month-$Day-$Year : $Hour : $Minute : $Second\n", TIME_DBG);
$tttime = $Hour * 3600 + $Minute * 60 + $Second;

	while ($in = <STCI_INPUT>) {
		if($in =~ /^\s*###/){
            next;
        } elsif($in =~ /^stcI_HASH\s*\:\s*(\S+)\s*$/){
			$stcI_for = $1;
		} elsif($in =~ /^stcI_EXTENSION\s*\:\s*(\S+)\s*$/){
			$stcI_extension = $1;
		} elsif($in =~ /^stcI_FILEPOSTFIX\s*\:\s*(\S+)\s*$/){
			$stcI_filepostfix = $1;
		} elsif($in =~ /^stcI_FILEBODY\s*\:\s*(\S+)\s*$/){
			$stcI_filebody = $1;
		} elsif($in =~ /^stcI_FILE_LOWER\s*\:\s*(\S+)\s*$/){
			$stcI_file_lower = $1;
		} elsif($in =~ /^stcI_ALL_LOWER\s*\:\s*(\S+)\s*$/){
			$stcI_all_lower = $1;
		} elsif($in =~ /^stcI_FILEPREFIX\s*\:\s*([^\s]+)\s*$/){
			$stcI_fileprefix = $1;
		} else {
			$stcI_lines .= $in;
		}
	}
	print_fp("stcI_HASH : $stcI_for\n",STDOUT,DBG);
	print_fp("stcI_EXTENSION : $stcI_extension\n",STDOUT,DBG);
	print_fp("stcI_FILEPOSTFIX : $stcI_filepostfix\n",STDOUT,DBG);
	print_fp("stcI_FILEPREFIX : $stcI_fileprefix\n",STDOUT,DBG);
	foreach $temp (keys %$stcI_for){
        my $mykeyfilename = $temp;
        if($stcI_file_lower eq "YES"){ $mykeyfilename = lc($temp); }
		print_fp("stcI_HASH : $stcI_for key $temp\n",STDOUT,DBG);
		if($stcI_filebody eq "NO"){
            $stcfilename = "$stcI_fileprefix" . "$stcI_filepostfix";
        } else {
            $stcfilename = "$stcI_fileprefix" . "$mykeyfilename" . "$stcI_filepostfix";
        }
		$stcfilename =~ s/KEY/$temp/g;
		$stcfilename =~ s/VALUE/$$stcI_for{$temp}/g;
        print "final file name 1 : $stcfilename";
		$stcfilename = replace_var_with_value($stcfilename);
        print "final file name 2 : $stcfilename";
		$ttt = replace_var_with_value($ttt);
        if($stcI_all_lower eq "YES"){ $stcfilename = lc($stcfilename); }
        my $stcNoDirectoryName = $stcfilename;
        $stcNoDirectoryName =~ s/\//_/g;
		print_fp("STCI tmp OUTPUT FileName :  $stcNoDirectoryName\n",STDOUT,DBG);
		print_fp("STCI OUTPUT FileName :  $stcfilename\n",STDOUT,DBG);
		my $tmpStcIFileName = "$outputdir/tmp/$stcNoDirectoryName\.$stcI_extension\.stc";
		open(STCI_OUTPUT , ">$tmpStcIFileName");
		print STCI_OUTPUT "FileName : $stcfilename\.$stcI_extension\n";
		print "FileName : $stcfilename\.$stcI_extension\n";
		@stcI_lines = split("\n",$stcI_lines);
		foreach my $ttt (@stcI_lines){
#print "ttt $ttt\n";
			if($ttt =~ /^SetI\s*\:/){
#print "A0 $ttt\n";
				$ttt =~ s/KEY/$temp/g;
				$ttt =~ s/VALUE/$$stcI_for{$temp}/g;
#print "A1 $ttt\n";
				$ttt = replace_var_with_value($ttt);
				$ttt =~ s/^SetI/Set/;
#print "A2 $ttt\n";
				print STCI_OUTPUT "$ttt\n";
			} else {
				print STCI_OUTPUT "$ttt\n";
			}
		}
		close(STCI_OUTPUT);
		CChange("$tmpStcIFileName","stc","","DEBUG_ON");			# OUTPUT¾È¿¡ ÀúÀåÇÏ°í, makefile¿¡µµ ³Ö¾î¶ó.
	}
	close(STCI_INPUT);

($Second, $Minute, $Hour, $Day, $Month, $Year, $WeekDay, $DayOfYear, $IsDST) = localtime(time) ;
$Month++;
$Year += 1900;
$ddtime = ($Hour * 3600 + $Minute * 60 + $Second) - $tttime;
print_fp( "structg Ichange duration ($ddtime) : Ichange END ($stcI_filename_input) - $Month-$Day-$Year : $Hour : $Minute : $Second\n",TIME_DBG);

}

sub CChange
{
	my $iterate_key;
	my $iterate_value;
	my $iterate_var_name;
	my $iterate_var_type;
	my $iterate_lines = "";
	my %file_output;
	my $in;
	my $in_start;
	my $in_end;
	my $iterate_cnt = 0;
	my $stc_filename_input ;
	my $stc_filename_output ;
	my $stc_debug ;
	my $stc_output_dir;
	my $tttime;
	my $ddtime;
	my $cchange_start_time;
	my $myiterateNewLineType = "";

	$stc_filename_input = shift @_;
	$stc_output_dir = shift @_;
	$stc_filename_output = shift @_;
	$stc_debug = shift @_;
	print "STC input $stc_filename_input , dir= $stc_output_dir , output= $stc_filename_output , debug= $stc_debug\n";
	#
	#방법 1. 
	# 쌍을 알기위해서 recusive를 사용하게 되고,
	# 그 안에서 나온 결과를 이용하여
	# 실제 들아갈 값을 찾는 것은 loop를 계속 돌며, ITERATOR가 없을때까지 수행을 하면 된다. 
	#  	loop를 돌때는 ITERATOR가 한번 발견되면 그 규칙을 끝까지 지키는 것을 원칙으로 한다.
	# 그리고, 나머지 loop들도 돌면서 처리를 계속한다. 
	#
	# recusive모양은 찾는 것은 좋으나, 거꾸로 풀면서 가면 위에서 지정한 loop의 argument를 사용할수 없기에 loop를 더 추가한 것이다.
	#
	#방법 2.
	# ITERATOR에 대한 변수를 두어 처음부터 잡아서 해당 것의 끝까지 저장을 한다.
	# ITERATOR A start
	# 	AAA
	# 	ITERATOR B start
	# 		BBB
	#   B End
	# A End
	# 라고 할때, 위의 4줄을 모두 한 변수에 집어 넣는다. (AAA ~ B End)
	# ITERATOR A의 규칙대로 대치를 시킨다. (복수의 line을 뽑아내어 변수에 저장)
	# 대치시킨 결과를 CChange() 에서 하는 것 처럼 또 돌려준다. 
	# ITERATOR B ~ B End까지를 또 풀어준다.
	#
	# operation 우선순위
	#  ITERATE 가 제일 먼저 처리 된다.
	#  그 후에 , +<+$...+>+ 이 처리가 된다.  co_log_main.c안에서  Get_HashoData를 참조하시요.
	print_fp("CChange $stc_filename_input\n",STDOUT,DBG);
	open(INPUTC , "<$stc_filename_input");

	if($stc_debug eq "DEBUG_ON"){ start_time_log("==time_debug=="); }

	# for make file : but not use yet
my $comment =  <<END_COMMENT;
	if($stc_filename_output ne ""){
		$stg_stc_file{$stc_filename_output} = $stc_filename_output;
		print_fp("STC FileName 1 : stg_stc_file : $stg_stc_file{$stc_filename_output} : $stc_filename_output\n",STDOUT,DBG);
		open(OUTPUTC , ">$outputdir/$stc_output_dir/$stc_filename_output");
		if( ($stc_filename_output =~ /\.c\s*$/) || ($stc_filename_output =~ /\.cpp\s*$/) || ($stc_filename_output =~ /\.pc\s*$/) ){	
			#print "INCLUDE1\n";
			print OUTPUTC "\n#include \"$FileName\"\n\n\n";		# .c안에서는 되는데 .l에서 문제가 발생함.
		} else {
			#NONE#
			# .l일때는 처음에 include를 선언하면 error가 발생한다.
		}
		if($stc_output_dir eq ""){
			if($stc_filename_output =~ /\.c/){
				$filelist{$stc_filename_output} = "CFILE";
			}
		}
	}
END_COMMENT


	while ($in = <INPUTC>) {
		if($in =~ /^FileName\s*\:\s*(\S+)\s*$/){
			$stc_filename_output = $1;
			$stg_stc_file{$1} = $1;
			print_fp("STC FileName 2 : stg_stc_file : $stg_stc_file{$1} : $stc_filename_output\n",STDOUT,DBG);
			$file_output{$stc_filename_output} = "";        # init

my $comment =  <<END_COMMENT;		# if need this statements , you change to put the values into hash variables.
			if( ($stc_filename_output =~ /\.c\s*$/) || ($stc_filename_output =~ /\.cpp\s*$/) || ($stc_filename_output =~ /\.pc\s*$/) ){	
				#print "INCLUDE2\n";
				print OUTPUTC "\n#include \"$FileName\"\n\n\n";		# .c안에서는 되는데 .l에서 문제가 발생함.
			} else {
				#NONE#
				# .l일때는 처음에 include를 선언하면 error가 발생한다.
			}
END_COMMENT

			# HASH : filelist
			my $tmpDir;
			if($stc_output_dir eq ""){
				$tmpDir = "";
			} else {
				$tmpDir = "$stc_output_dir\/";
			}
			if($stc_filename_output =~ /\.c/){
				$filelist{$tmpDir.$stc_filename_output} = "CFILE";
			} elsif($stc_filename_output =~ /\.l/){
				$filelist{$tmpDir.$stc_filename_output} = "LEXFILE";
				$temp = $stc_filename_output;
				if($temp =~ /(.*)\/(.*).l/){
					$temp = "lex\.$2\.c";
				} else {
					$temp =~ s/(.*)\.l/lex\.$1\.c/;
				}
				$filelist{$tmpDir.$temp} = "CFILE";
			}
			$hashName{filelist} = $stc_filename_output;
			next;
		} elsif($in =~ /^Set\s*\:\s*\$(\w+)\s*\=\s*(\S+)/){		# Set: $A = B 
			my $temp_set_var = $1;
			my $temp_set_value = $2;
print "LLL $iterate_comments : [$1]  [$2]\n";
			$$1 = $2;
#print DBG "LLL $iterate_comments : [$1] [$2] \n";
			#CCHANGE_DEBUG print DBG "SET : $1 = $2\n";
			#  이게 SET하는 것이 잘 안되는군.... 다음 방법은 될까?  <- set을 잘못하여 시험함.
			# 위의 방법 1과 아래의 방법 2 모두 잘 처리됨.
##			if($temp_set_var eq "iterate_comments" ){
##				print "iterate_comments : [$1] [$2]\n";
##				if("OFF" eq $temp_set_value){
##					#$iterate_comments = "OFF";
##				} else {
##					#$iterate_comments = "ON";
##				}
##			}
##			# 첫번째의 위방법은 이상하게 한 번만 Set되고 되지 않는다. 
##			# 두번째인 비교를 하여 직접 값을 써주는 것은 가능하다. ㅁ
			$undefined_name = "stc_Set_var_$stc_filename_output";
			$undefined_typedef{$undefined_name} = "HASH_$undefined_name";
			$$undefined_name{$temp_set_var} = "$temp_set_value";
			$local_var_set{$stc_filename_output}{$temp_set_var} = $temp_set_value;
			$hashName{local_var_set} = "STC Local Set Variable";
			print "$stc_filename_output : $in\n";
			next;
		} elsif($in =~ /^Set\s*\:\s*\$(\w+)\s*\{\s*(\w+)\s*\}\s*\=\s*(\w+)/){		# Set: $A{B} = C
			print "SET : ERROR $1 $2 $3\n";
			my $temp_set_var = $1;
			my $temp_set_hash = $2;
			my $temp_set_value = $3;
#print DBG "LLL $iterate_comments : $1  $2 $3\n";
			#  $hashName is the list of hash variables to print or to write into file.
			$hashName{$1} = $3;
			$$1{$2} = $3;
#print DBG "LLL $iterate_comments : [$1] [$2]  $3\n";
			#CCHANGE_DEBUG print DBG "SET :  $1 \{ $2 \} = $3\n";
			#$undefined_name = "stc_Set_var_$stc_filename_output";
			#$undefined_typedef{$undefined_name} = "HASH_$undefined_name";
			#$$undefined_name{"$temp_set_var" . "$temp_set_hash"} = "$temp_set_value";
			next;
		}


#print_fp("Line 1 icnt=$iterate_cnt : $in",DBG);
		if(0 == $iterate_cnt){
	            #foreach my $key (sort keys %API) { print DBG __SUB__ . "API1-4-0 $key - CChange\n"; }
			#$in = replace_var_with_value($in,3696);
	            #foreach my $key (sort keys %API) { print DBG __SUB__ . "API1-4-1 $key - CChange\n"; }
		}
#print_fp("Line 2 icnt=$iterate_cnt : $in",DBG);

		if ($in =~ /^\s*ITERATE(WithoutNewLine\s+|\s+)([+-]?[KV]?[%@&+])(\S+)\s+\+<<\+\s+(\S+)\s+(\S+)\s*$/){
			#ITERATOR_DEBUG 
			print DBG "ITERATE Mstart $1 $2 $3 $4 $5\n"; 
			if(0 == $iterate_cnt){
				$in_start = $in;
				($iterate_var_type , $iterate_var_name , $iterate_key , $iterate_value) = ($2,$3,$4 , $5);
				$myiterateNewLineType = $1;
			} else {
				$iterate_lines .= $in;
			}
			$myiterateNewLineType =~ s/\s*$//;
			#print STDERR "[$myiterateNewLineType]\n";
			$iterate_cnt ++;
			#ITERATOR_DEBUG 
			print DBG "ITERATE iterate_cnt $iterate_cnt : $in";
		} elsif ($in =~ /^(.*)\+>>\+/){
			$iterate_cnt--;
			#ITERATOR_DEBUG 
			print DBG "ITERATE iterate_cnt $iterate_cnt : $in";
			if(0 == $iterate_cnt){
				if($stc_debug eq "DEBUG_ON"){ mid_time_log("==MID time_debug=="); }
				$in_end = $in;
				#ITERATOR_DEBUG 
				print DBG "iterate_comments [" . $iterate_comments . "]\n";
                my $oDebug = "";
				if( $optionDebug ){
					$temp1=$in_start;
					$temp2=$iterate_lines;
					$temp3=$in_end;
					# /** */ 으로 묶는 안에 /* */이 있으면 안되므로 <* *>으로 변환을 시켜주는 것이다.  
					#$temp1 =~ s/\/\*/\<\*/g;
					#$temp1 =~ s/\*\//\*\>/g;
					#$temp2 =~ s/\/\*/\<\*/g;
					#$temp2 =~ s/\*\//\*\>/g;
					$temp2 =~ s/IFNOTEQUAL/ifNOTequal/g;
					$temp2 =~ s/IFEQUAL/ifequal/g;
					#$temp3 =~ s/\/\*/\<\*/g;
					#$temp3 =~ s/\*\//\*\>/g;
					#ITERATOR_DEBUG 
					print DBG "\/\*\*\n$temp1$temp2$temp3\*\/\n";
                    $oDebug = "#if 0 // DEBUG ITERATE  : debug mode start\n$temp1$temp2$temp3\n#endif // debug mode end\n";
					$oDebug =~ s/\+\{\{\+/\+\[\[\+/g;
					$oDebug =~ s/\+\}\}\+/\+\]\]\+/g;
					$oDebug =~ s/\+\<\+/\+\(\+/g;
					$oDebug =~ s/\+\>\+/\+\)\+/g;
					$oDebug =~ s/\+\<\<\+/\+\(\(\+/g;
					$oDebug =~ s/\+\>\>\+/\+\)\)\+/g;
				}
				# Iterator_recursion은 단지 확장을 위한 것이다. 그러므로 , 확장을하는 것만 해주면 된다.  
				$iterate_lines = Iterator_recursion($iterate_var_type , $iterate_var_name,$iterate_key,$iterate_value,$iterate_lines);

				if( $optionDebug ){
                    $iterate_lines = $oDebug . $iterate_lines;
                }

				#ITERATOR_DEBUG  
				#$iterate_lines =~ s/\+<\+\s*\$(\S+)\s*\+>\+/$$1/g;		# 	+<+$stg_hash_del_timeout+>+ ==> 10

				if($stc_debug eq "DEBUG_ON"){ mid_time_log("==MID time_debug=="); }

				# 이런식으로 처리하면 많은 %값들을 만들지 않아도 되며, define같은 값들을 지저분하게 군데군데 만들어줄 필요가 없다. 
				#print DBG "Set Hash 10 : $iterate_lines\n";
				if(0){      # It is mendatory
                    # for performance
					#  because of processing speed.  when it is replacement , it scans whole string. So I break down into substring.
					my $iter_lena = length($iterate_lines);
					my $iterate_lines_org = $iterate_lines;
					$iterate_lines ="";
					for(my $itt = 0;$itt <= $iter_lena ; $itt += 10000){
						$iterate_lines .= replace_var_with_value(substr($iterate_lines_org, $itt, 10000));
					}

					$temp = length($iterate_lines);
					$iterate_lines = replace_var_with_value($iterate_lines);

					if($stc_debug eq "DEBUG_ON"){ mid_time_log("==MID time_debug 2=="); }

				} else {            # This is old version's code.
                    #$iterate_lines = replace_var_with_value($iterate_lines);
				}
				#print DBG "Set Hash 11 : $iterate_lines\n";
				print DBG "RETURN \$iterate_lines = \n\[\n$iterate_lines\n\]\n";

				$iterate_lines = recover_special_code($iterate_lines);
				#print_fp("$iterate_lines\n" , OUTPUTC);
				#print STDERR "1[$myiterateNewLineType]\n";
				if($myiterateNewLineType ne "WithoutNewLine"){
					$file_output{$stc_filename_output} .= $iterate_lines;
				} else {
					#print STDERR "1{$iterate_lines}\n\n";
					#print STDERR "1[ $file_output{$stc_filename_output} ]\n\n";
					$file_output{$stc_filename_output} =~ s/\n$/ /;
					#print STDERR "1[ $file_output{$stc_filename_output} ]\n\n";
					$iterate_lines =~ s/\n/ /g;
					$iterate_lines =~ s// /g;
					$file_output{$stc_filename_output} .= $iterate_lines . "\n";
					$myiterateNewLineType = "";
					#print STDERR "2{$iterate_lines}\n\n";
					#print STDERR "2[ $file_output{$stc_filename_output} ]\n\n";
				}

				$iterate_lines = "";
#print DBG "Mend \n";
			} else {			# if(0 == $iterate_cnt)
				$iterate_lines .= $in;
			} 					#if(0 == $iterate_cnt)
		} else {
			if(0 == $iterate_cnt){
				#print_fp($in, OUTPUTC);
				$file_output{$stc_filename_output} .= $in;
			} else {
				$iterate_lines .= $in;
			}
		}
	}

	print DBG __SUB__ . " iterate_equal\n";
	foreach my $tmpKey  (sort keys  %file_output){
		if($stc_debug eq "DEBUG_ON"){ mid_time_log("==MID equal start =="); }
		my $iter_len = length($file_output{$tmpKey});
		my $linesOrg = $file_output{$tmpKey};
		my $lines ="";
        #print DBG __SUB__ . " CCCCC IFEQUAL = $linesOrg\n";
		if(0){
			# for performance
			for(my $itt = 0;$itt <= $iter_len ; $itt += 1000){
				$lines .= iterate_equal(substr($linesOrg, $itt, 1000));
			}
			$lines = iterate_equal($lines);
		} else {
            if($optionPerformance){
			    $lines = iterate_equal_performance($linesOrg);
            } else {
			    $lines = iterate_equal($linesOrg);
            }
		}
        #print DBG __SUB__ . " DDDDD IFEQUAL = $linesOrg\n";
		$lines = replace_var_with_value($lines);
        $lines = recover_special_code($lines);
		print DBG "FFFF $lines\n";
		#$lines =~ s/STG_SHARP_/\#/g;
		if($stc_debug eq "DEBUG_ON"){ mid_time_log("==MID equal end =="); }

        #print "tmpKey $tmpKey\n";
        if($tmpKey =~ /[^\/]*$/){
			my $m_before = $`;
			my $m_match = $&;
            #print "tmpKey before[$m_before] , match [$m_match]\n";
            if($m_before ne ""){
                if($tmpKey =~ /^\//){
                    print "make_path $m_before\n";
                    make_path("$m_before");
                } else {
                    print "make_path $outputdir/$stc_output_dir/$m_before\n";
                    make_path("$outputdir/$stc_output_dir/$m_before");
                }
            }
        }
        if($tmpKey =~ /^\//){
		    open(OUTPUTC , "> $tmpKey");
        } else {
		    open(OUTPUTC , "> $outputdir/$stc_output_dir/$tmpKey");
        }
		$lines =~ s///g;
		print OUTPUTC $lines;
		close(OUTPUTC);
	}

	close(INPUTC);

	if($stc_debug eq "DEBUG_ON"){ end_time_log("==END CChange =="); }
}

sub iterate_equal_performance()
{
	# Rules
	# IFEQUAL | IFNOTEQUAL ( condition with general rules :&& ,|| ,etc)  /#       
	#       Contents with multiple lines
	# #/
	my $il="";
	my $ifequal_one="";
	my $ifequal_two="";
	my $ifequal_parm="";
	my $len;
	my $if_before="";
	my $if_match="";
	my $if_after="";
    my $if_eval = "";

	$il = shift @_;


	open(STF , ">temp.cpp.stc");
    print STF $il;
    close(STF);

    my $cnt = 0;

    
    my @ll = split(/\n/,$il);
    print DBG $ll[0] . "\n";
    print DBG $ll[1] . "\n";
    my $startll = @ll;
    my $sbrace;
    my $ebrace;
    my @stack;
    my @oDbg;
    print $startll . "\n";
    while($startll >= 0){
        print DBG "While startll $startll\n";
        if($ll[$startll] =~ m/\+\}\}\+/){
            push @stack , $startll;
            print DBG " stack:@stack   , count $#stack\n";
            $startll--;
        } elsif($ll[$startll] =~ m/\+\{\{\+/){
            $sbrace = $startll;
            print DBG "stack:" . "@stack \n";
            $ebrace = pop @stack;
            print DBG "startll $startll  , sbrace $sbrace , ebrace $ebrace\n";
            print DBG "stack:" . "$#stack : @stack \n";
            if($ll[$startll] =~ m/(IFEQUAL|IFNOTEQUAL)/){
                my $order = $&;
                my $after = $';
                my $condition;
                my $val;
                if($after =~ m/(\(([^\(\)]|(?R))*\))/){
                    $condition = $&;
                    print DBG "$startll : $order  condition = $condition\n";
                }
		        $condition = replace_var_with_value($condition);
                $val = eval($condition);
                print DBG "$startll : $order  condition = $condition , value:$val\n";
                if($order eq "IFEQUAL"){
                    if($val ne ""){
                        print DBG "==1== , stacksize $#stack\n";
                        for($i=10;$i>=-10;$i--){
                            if( ($sbrace <= ($startll-$i) ) && (($startll-$i) <= $ebrace) ){
                                print DBG "o| ";
                            } else {
                                print DBG ">> ";
                            }
                            print DBG ($startll - $i) . " : $ll[$startll-$i]\n";
                        }
                        if($optionDebug){
                            splice(@oDbg , $ebrace , 1);
                            splice(@oDbg , $sbrace , 1);
                            if($#stack < 0){
                                $oDbg[$sbrace] .= "#if 0 // DEBUG_IF : debug mode\n";
                                my $j=0;
                                for($i=$sbrace;$i<=$ebrace;$i++){
                                    $oDbg[$sbrace] .= $ll[$i] . "\n";
                                    $j++;
                                    if($j > 8){
                                        $oDbg[$sbrace] .= "....\n";
                                        last;
                                    }
                                }
                                $oDbg[$sbrace] .= "#endif // DEBUG_IF : debug mode\n";
                            }
                        }
                        splice(@ll , $ebrace , 1);
                        splice(@ll , $sbrace , 1);
                        for($i=10;$i>=-10;$i--){
                            if( ($sbrace <= ($startll-$i) ) && (($startll-$i) <= ($ebrace-2)) ){
                                print DBG "o| ";
                            } else {
                                print DBG "<< ";
                            }
                            print DBG ($startll - $i) . " : $ll[$startll-$i]\n";
                        }
                        for($i=0;$i<=$#stack;$i++){
                            print DBG "old stack : [$i] $stack[$i] -> ";
                            $stack[$i] = $stack[$i] - 2;
                            print DBG "new stack : [$i] $stack[$i]\n";
                        }
                        $startll = $startll -1;
                        print DBG "mid startll : $startll\n";
                    } else {
                        print DBG "==2==\n";
                        for($i=10;$i>=-10;$i--){
                            if( ($sbrace <= ($startll-$i) ) && (($startll-$i) <= $ebrace) ){
                                print DBG "x| ";
                            } else {
                                print DBG ">> ";
                            }
                            print DBG ($startll - $i) . " : $ll[$startll-$i]\n";
                        }
                        if($optionDebug){
                            splice(@oDbg , $sbrace , $ebrace - $sbrace +1);
                        }
                        splice(@ll , $sbrace , $ebrace - $sbrace +1);
                        for($i=10;$i>=-10;$i--){
                            if( ($sbrace == ($startll-$i) ) ){
                                print DBG "x| ";
                            } else {
                                print DBG "<< ";
                            }
                            print DBG ($startll - $i) . " : $ll[$startll-$i]\n";
                        }
                        for($i=0;$i<=$#stack;$i++){
                            print DBG "old stack : [$i] $stack[$i] -> ";
                            $stack[$i] = $stack[$i] - ($ebrace - $sbrace +1);
                            print DBG "new stack : [$i] $stack[$i]\n";
                        }
                        $startll = $startll - 1 ;
                        print DBG "mid startll : $startll\n";
                    }
                } elsif($order eq "IFNOTEQUAL"){
                    if($val eq ""){
                        if($optionDebug){
                            splice(@oDbg , $ebrace , 1);
                            splice(@oDbg , $sbrace , 1);
                            if($#stack < 0){
                                $oDbg[$sbrace] .= "#if 0 // DEBUG_IF : debug mode\n";
                                my $j=0;
                                for($i=$sbrace;$i<=$ebrace;$i++){
                                    $oDbg[$sbrace] .= $ll[$i] . "\n";
                                    $j++;
                                    if($j > 8){
                                        $oDbg[$sbrace] .= "....\n";
                                        last;
                                    }
                                }
                                $oDbg[$sbrace] .= "#endif // DEBUG_IF : debug mode\n";
                            }
                        }
                        splice(@ll , $ebrace , 1);
                        splice(@ll , $sbrace , 1);
                        for($i=0;$i<=$#stack;$i++){
                            $stack[$i] = $stack[$i] - 2;
                        }
                        $startll = $startll -1;
                    } else {
                        if($optionDebug){
                            splice(@oDbg , $sbrace , $ebrace - $sbrace +1);
                        }
                        splice(@ll , $sbrace , $ebrace - $sbrace +1);
                        for($i=0;$i<=$#stack;$i++){
                            $stack[$i] = $stack[$i] - ($ebrace - $sbrace +1);
                        }
                        $startll = $startll - 1 ;
                    }
                }
            } else {
                print STDERR "(temp.cpp.stc) line:" . $startll+1 . " does not have IF command\n";
                exit(4);
            }
            if($#stack >= 0){
                $startll = pop @stack;
            } else {
                #$startll--;
            }
            #for($i=10;$i>=-5;$i--){
                #print DBG ">> " . ($startll - $i) . " : $ll[$startll-$i]\n";
            #}
            print DBG ">end stack count:" . "$#stack \n";
            print DBG "end stack : @stack \n";
            print DBG "end startll : $startll \n";
        } else {
            $startll--;
        }
    }

    #$il = join("\n",@ll);
    $il = "";
    for($i=0;$i<=$#ll;$i++){
        if($optionDebug){
            print DBG "optionDebug $i : $oDbg{$i}\n";
            if($oDbg[$i]){
                $il .= $oDbg[$i];
            }
        }
        $il .= $ll[$i] . "\n";
        print DBG "line $i : $ll[$i]\n";
    }

	open(STF , ">temp.cpp");
    print STF $il;
    close(STF);

    return $il;
} 

sub iterate_equal()
{
	# Rules
	# IFEQUAL | IFNOTEQUAL ( condition with general rules :&& ,|| ,etc)  /#       
	#       Contents with multiple lines
	# #/
	my $il="";
	my $ifequal_one="";
	my $ifequal_two="";
	my $ifequal_parm="";
	my $len;
	my $if_before="";
	my $if_match="";
	my $if_after="";
    my $if_eval = "";

	$il = shift @_;


	#open(STF , ">temp.cpp");
    #print STF $il;
    #close(STF);

    $tcnt++;
    print __SUB__ . "\n our iterate_equal $tcnt\n";

    
    my $cnt = 0;
    while ($il =~ m/\+\{\{\+(\s*(.*?|\n)*\s*)+\}\}\+/){   # recursion for nesting
        #print("cnt $cnt , end $end\n");

        $pttn = "\+\{\{\+";
        if($il =~ m/\+\{\{\+(\s*(.*?|\n)*\s*)+\}\}\+/){   # recursion for nesting
            $before = $`;
            $match = $&;
            $after = $';
            #print DBG __SUB__ . "ND1-1 match[$match]\n";
            @a = split(/\+\{\{\+/,$match);
            #print DBG __SUB__ . "ND1-1 \n";
            for($i=0;$i<= $#a;$i++) {
                #print DBG __SUB__ . "$i:($a[$i])\n";
            }
            $final = $a[$#a];
            delete $a[$#a];
            for($i=0;$i<= $#a;$i++) {
                #print DBG __SUB__ . "$i:($a[$i])\n";
            }

            $final =~ s/^\s*\n//;
            $final =~ s/(\n)?\s*\+\}\}\+\s*$//;
            #print DBG __SUB__ . "final:($final)\n";
            #print "before($before)\n";
            for($i=0;$i< $#a;$i++) {
                $before .= $a[$i] . $pttn;
            }
        }
        if($a[$#a] =~ m/(IFEQUAL|IFNOTEQUAL)(\s*)/){
            $i_before = $`;
            $i_match = $&;
            $i_after = $';
            $iif = $1;
            #print DBG __SUB__ . "ND1-2 i_before[$iif:$i_before]\n";
            #print DBG __SUB__ . "ND1-2 i_match[$i_match]\n";
            ##print DBG __SUB__ . "ND1-2 i_after[$i_after]\n";
            $i_before =~ s/\n\s*$//;
            $before .= $i_before;
        } else {
            if($before =~ m/(IFEQUAL|IFNOTEQUAL)(\s*)/){
                $i_before = $`;
                $i_match = $&;
                $i_after = $';
                $iif = $1;
                #print DBG __SUB__ . "ND1-3 i_before[$iif:$i_before]\n";
                #print DBG __SUB__ . "ND1-3 i_match[$i_match]\n";
                ##print DBG __SUB__ . "ND1-3 i_after[$i_after]\n";
                $i_before =~ s/\n\s*$//;
                $before = $i_before;
            } 
        }
        #print DBG __SUB__ . "before($before)\n";
        if($i_after =~ m/(\(([^\(\)]|(?R))*\))/){
            $b_before = $`;
            $b_match = $&;
            $b_after = $';
            #print DBG __SUB__ . "ND1-2 b_before[$b_before]\n";
            #print DBG __SUB__ . "ND1-2 b_match[$b_match]\n";
            ##print DBG __SUB__ . "ND1-2 b_after[$b_after]\n";
        }

		$b_match = replace_var_with_value($b_match);
        if($iif eq "IFEQUAL"){
            $val = eval($b_match);
            #print DBG __SUB__ . "$b_match -> val:$val\n";
            if($val ne ""){
                if($final ne ""){ $before .= "\n$final"; }
            }
        } else {
            $val = eval($b_match);
            #print DBG __SUB__ . "$b_match -> val:$val\n";
            if($val eq ""){
                if($final ne ""){ $before .= "\n$final"; }
            }
        }
        ##print "before($before$after)\n";
        #open($fh , ">" , "iter" . $cnt . ".out");
        ##print $fh "$before$after\n";
        #close $fh;
        $il = "$before$after";
        $cnt++;
        if($cnt == $end){ return ""; }
        #print DBG __SUB__ . "cnt $cnt , end $end\n";
        ##print DBG __SUB__ . " RD1 $il]]]]]\n";
    };

    #print "iterate_equal loop count $cnt end\n";

	return $il;
}

our %pNull = {};
sub getHashRef {
	my ($name) = @_;        # gCan{9}
	my $first;
	#print "F " . $name . "\n";
	$name =~ s/([^\{]+)//;
	$first = $1;
	my $hn = \%{$first};
	#print "F " . $hn . "\n";
	while($name =~ s/^\{([^\}]*)\}//){
		my $uu = $1;
		$hn = $hn->{$uu};
		#print "G " . $uu . "   $hn\n";
	}

	#print "I " . $hn . "\n";
	if($hn == 0){
		return \%pNull;
	} else {
		return $hn;
	}
}

sub max_keys
{
	my $iterate_var_name;
	my $allDigit = 1;
	my $max=0;
	($iterate_var_name ) = @_;
	foreach my $tmpKey ( keys %{getHashRef($iterate_var_name)}){
		if(not ($tmpKey =~ /^\s*\d*\s*$/)){
			$allDigit = 0;
			last;
		}
		if($max < $tmpKey){ $max = $tmpKey; }
	}
	
	if($allDigit == 1){ return  $max; }
	else { return  0; }
}

sub sort_keys
{
	my $iterate_var_name;
	my $iterate_op_type;
	my $allDigit = 1;
	my $tmpReverse = "+";
	my $tmpValue = 0;
	my $debug = "DEBUG_OFF";
	my @ret;
	($iterate_var_name ,$iterate_op_type,$debug) = @_;
	if($iterate_op_type =~ m/-/){ $tmpReverse = "-"; }
	if($iterate_op_type =~ m/V/){ $tmpValue = "V"; }
	if($tmpValue eq "V"){
		foreach my $tmpKey ( keys %{getHashRef($iterate_var_name)}){
			if(not (getHashRef($iterate_var_name)->{$tmpKey} =~ /^\s*\d*\s*$/)){
				$allDigit = 0;
				last;
			}
		}
	} else {
		foreach my $tmpKey ( keys %{getHashRef($iterate_var_name)}){
			if(not ($tmpKey =~ /^\s*\d*\s*$/)){
				$allDigit = 0;
				last;
			}
		}
	}
		
	if($allDigit == 1){ 
		if($tmpValue eq "V"){
			@ret = sort {getHashRef($iterate_var_name)->{$a} <=> getHashRef($iterate_var_name)->{$b}} keys %{getHashRef($iterate_var_name)};
		} else {
			@ret = sort {$a <=> $b} keys %{getHashRef($iterate_var_name)};
		}
	}
	else { @ret =  sort keys %{getHashRef($iterate_var_name)}; }

	if($tmpReverse eq "-"){
		@ret = reverse @ret;
	}

	if($debug eq "DEBUG_ON"){
		print DBG __SUB__;
		foreach (@ret){
			print DBG "  $_";
		}
		print DBG "\n";
	}

	return @ret
}

sub Iterator_recursion
{
	my $iterate_var_type;
	my $iterate_var_name;
	my $iterate_key;
	my $iterate_value;
	my $iterate_lines;
	my $result = "";
	my $in;
	my $iterate_cnt = 0;
	my @lines;
	my $myiterate_newline_type = "";

	($iterate_var_type , $iterate_var_name , $iterate_key , $iterate_value, $iterate_lines) = @_;
	#print_fp( "O : @_\n", OUTPUTC);
	#print DBG __SUB__ . "RC1 : $iterate_var_type , $iterate_var_name , $iterate_key , $iterate_value\n";

	#print DBG __SUB__ . "RC2 : Iterator_recursion : \$iterate_lines = $iterate_lines ]]]\n";
	$iterate_var_name = replace_var_with_value($iterate_var_name);

	# Various Operation  :
	#   % : hash
	#   @ : array
	#   & : reverse array
	#   + : 0 ~ max key value (only digits)
	if($iterate_var_type =~ m/^[+-]?[KV]?\%/){
		print DBG __SUB__ . "RC : HASH Iterator_recursion : \$iterate_var_name = $iterate_var_name ]]]\n";
		#$tmp1 = eval $$iterate_var_name;
		#print DBG __SUB__ . "RC-1 : " . (sort_keys($iterate_var_name) ) . "\n";
		foreach $stg_key_hash (sort_keys($iterate_var_name,$iterate_var_type,"DEBUG_ON")){
			print DBG __SUB__ . "RC-2 : HASH2 Iterator_recursion : \$key = $stg_key_hash value=[ $$iterate_var_name{$stg_key_hash} ] [ $iterate_var_name ] [ " .  getHashRef($iterate_var_name)->{$stg_key_hash} . "]\n";
			my $myvalue = getHashRef($iterate_var_name)->{$stg_key_hash};
			$temp = $iterate_lines;
			$temp =~ s/$iterate_key/$stg_key_hash/g;
			$temp =~ s/$iterate_value/$myvalue/g;
			$result .= $temp;
		}
	} elsif($iterate_var_type eq "\@"){
		my $my_cnt;
		$my_cnt = @$iterate_var_name;
		#ITERATOR_DEBUG print DBG "--> ARRAY : $iterate_var_name  size =  $my_cnt\n";
		for(my $i = 0 ; $i < $my_cnt ; $i++){
			#ITERATOR_DEBUG print DBG "array : \$$iterate_var_name \[ $i \] = $$iterate_var_name[$i]\n";
			$temp = $iterate_lines;
			$temp =~ s/$iterate_key/$i/g;
			$temp =~ s/$iterate_value/$$iterate_var_name[$i]/g;
			$result .= $temp;
		}
	} elsif($iterate_var_type eq "\&"){
		my $my_cnt;
		$my_cnt = @$iterate_var_name;
		#ITERATOR_DEBUG print DBG "--> REVERSE ARRAY : $iterate_var_name  size =  $my_cnt\n";
		for(my $i = $my_cnt - 1 ; $i >= 0 ; $i--){
			#ITERATOR_DEBUG print DBG "REVERSE array : \$$iterate_var_name \[ $i \] = $$iterate_var_name[$i]\n";
			$temp = $iterate_lines;
			$temp =~ s/$iterate_key/$i/g;
			$temp =~ s/$iterate_value/$$iterate_var_name[$i]/g;
			$result .= $temp;
		}
	} elsif($iterate_var_type eq "\+"){
		my $my_max = max_keys($iterate_var_name);
		#ITERATOR_DEBUG 
		print DBG "--> MAX : $iterate_var_name  MAX COUNT =  $my_max\n";
		for(my $i = 0 ; $i <= $my_max ; $i++){
			#ITERATOR_DEBUG print DBG "array : \$$iterate_var_name \[ $i \] = $$iterate_var_name[$i]\n";
			$temp = $iterate_lines;
			$temp =~ s/$iterate_key/$i/g;
			$temp =~ s/$iterate_value/$$iterate_var_name{$i}/g;
			$result .= $temp;
		}
	} else {
		print "ERROR : unknown iterate var type  : $iterate_var_type\n";
		die $error = 500;
	}
	#print DBG __SUB__ . "RC : Iterator_recursion : \$result = $result ]]]\n";


	# Various Operation
	$iterate_lines = "";
	if($result =~ /\s*ITERATE(WithoutNewLine\s+|\s+)([+-]?[KV]?[%@&+])(\S+)\s+\+<<\+\s+(\S+)\s+(\S+)/){ 
		@lines = split("\n",$result);
		$result = "";
		foreach my $it_line (@lines){
			if ($it_line =~ /^\s*ITERATE(WithoutNewLine\s+|\s+)([+-]?[KV]?[%@&+])(\S+)\s+\+<<\+\s+(\S+)\s+(\S+)/){  
#print DBG "Set Hash 20 : $iterate_lines \n";
#print DBG "Set Hash 21 : $iterate_cnt : $it_line\n";
				$it_line =~ /^\s*ITERATE(WithoutNewLine\s+|\s+)([+-]?[KV]?[%@&+])(\S+)\s+\+<<\+\s+(\S+)\s+(\S+)/;  
				if(0 == $iterate_cnt){
#print  DBG "Sstart $1 $2 $3 $4\n"; 
					($iterate_var_type , $iterate_var_name , $iterate_key , $iterate_value) = ($2,$3,$4,$5);
					$myiterate_newline_type = $1;
				} else {
					$iterate_lines .= $it_line . "\n";
				}
				$iterate_cnt ++;
				$myiterate_newline_type =~ s/\s*$//;
			}
			elsif ($it_line =~ /^(.*)\+>>\+/){
#print DBG "SUB_ITERATE : $iterate_cnt : $it_line\n";
				$iterate_cnt--;

				if(0 == $iterate_cnt){
                    my $oDebug="";
                    if($optionDebugDetail){
                        if(not($iterate_lines =~ /^[\s\n]*$/)){
                            $oDebug = $iterate_lines;
					        #$temp2 =~ s/\/\*/\<\*/g;
					        #$temp2 =~ s/\*\//\*\>/g;
					        $oDebug =~ s/IFNOTEQUAL/ifNOTequal/g;
					        $oDebug =~ s/IFEQUAL/ifequal/g;
					        $oDebug =~ s/\+\{\{\+/\+\[\[\+/g;
					        $oDebug =~ s/\+\}\}\+/\+\]\]\+/g;
					        $oDebug =~ s/\+\<\+/\+\(\+/g;
					        $oDebug =~ s/\+\>\+/\+\)\+/g;
					        $oDebug =~ s/\+\<\<\+/\+\(\(\+/g;
					        $oDebug =~ s/\+\>\>\+/\+\)\)\+/g;
                            $oDebug = "#if  0  // DEBUG_DETAIL : debug mode in detail start\n$oDebug\n#endif // debug mode in detail end\n";
                        }
                    }
					$iterate_lines = Iterator_recursion($iterate_var_type , $iterate_var_name,$iterate_key,$iterate_value,$iterate_lines);
#print  DBG "send result 30 :: $iterate_lines\n"; 
                    if($optionDebugDetail){
                        $iterate_lines = $oDebug . $iterate_lines;
                    }
					#$iterate_lines = replace_var_with_value($iterate_lines);
					if($myiterate_newline_type ne "WithoutNewLine"){
						$result .= $iterate_lines;
					} else {
						#print STDERR "3{$iterate_lines}\n\n";
						#print STDERR "3[$result]\n\n";
						chomp($result);
						$iterate_lines =~ s/\n/ /g;
						$iterate_lines =~ s// /g;
						$result .= $iterate_lines . "\n";
						$myiterate_newline_type = "";
						#print STDERR "4{$iterate_lines}\n\n";
						#print STDERR "4[$result]\n\n";
					}
					#print OUTPUTC "SUB RETURN \$iterate_lines = \n\[\n$iterate_lines\n\]\n";
#print  DBG "Send result 31 :: $result\n"; 
					$iterate_lines = "";
				} else {
					$iterate_lines .= $it_line . "\n";
				}
			} 
			else {
#print DBG "SUB_ITERATE 40 : $iterate_cnt : $it_line\n";
				if(0 == $iterate_cnt){
                    #$it_line = replace_var_with_value($it_line);
#print DBG "SUB_ITERATE 41 : $iterate_cnt : $it_line\n";
					$result .= $it_line . "\n";
				} else {
					$iterate_lines .= $it_line . "\n";
#print DBG "TTT : \$iterate_lines = $iterate_lines ]]]\n";
				}
			} 
		}
	}
	return $result;
}


sub replace_var_with_value
{
	my $replace_in;
	my $in_cnt = 0;
	my $nn_cnt = 0;
	my $replace_in = shift @_;
	my $num = shift @_;

    my $ln = length($replace_in);

    if($num == 369){
	    print DBG __SUB__ . ":" . __LINE__ . " ---- : num $num  $replace_in\n";
	    foreach my $key (sort keys %API) { print DBG __SUB__ . "API1 $key - replace_var_with_value\n"; }
    }

    if($ln > 10000){
        print "length = " . length($replace_in) . "\n";
    }

	#print DBG __SUB__ . ":" . __LINE__ . " ---- : in_cnt $in_cnt\n";
	#while(
	#($replace_in =~ s/\+<\+\s*\$([\w\d\.]+)\s*\+>\+/$$1/)		# 	+<+$stg_hash_del_timeout+>+ ==> 10
	#|| ($replace_in =~ s/\+<\+\s*\$([\w\d\.]+)\s*\[\s*(\d*)\s*\]\s*\+>\+/$$1[$2]/)	# +<+$typedef_name[54]+>+  ==> COMBI_Accum
	#|| ($replace_in =~ s/\+<\+\s*\$([\w\d\.]+)\s*\{\s*([^\}]+)\s*\}\s*\+>\+/$$1{"$2"}/) 	# +<+$HASH_KEY_TYPE{uiIP}+>+ ==> IP4
	#){
	#print DBG __SUB__ . __LINE__ . " BBBB : 1 $replace_in\n";
	#$in_cnt ++;
	#print DBG "Set Hash replace2 in_cnt=$in_cnt: $replace_in \n";
	#}			# +<+$type{+<+$HASH_KEY_TYPE{uiIP}+>+}+>+  ==> int

	#####  while($replace_in =~ /\+<\+\s*(\$[\w\d\.\)\(]+\s*[^\+>]*)\+>\+/)		# 	+<+$stg_hash_del_timeout+>+ ==> 10
    if(0){      # when I use pulus($g_y) , it is not proper method. so I won't use it
        # for performance
		my $iter_lena = length($replace_in);
		my $replace_org  = $replace_in;
        my $replace_tmp;
        $replace_in = "";
		for(my $itt = 0;$itt <= $iter_lena ; $itt += 1000){
		    $replace_tmp = substr($replace_org, $itt, 1000);
            #if($ln > 10000){
            #print "[[[$replace_tmp]]]\n";
            #}
	        while($replace_tmp =~ /\+<\+\s*([^\+>]*)\s*\+>\+/)		# 	+<+$stg_hash_del_timeout+>+ ==> 10
	        {
                #my $match = $&;
		        my $val = eval($1);
                #print DBG __SUB__ . "REPLACE:" . " $match  $1 => value :  $val  , iterate_cnt : $iterate_cnt\n";
		        $replace_tmp =~ s/\+<\+\s*([^\+>]*)\s*\+>\+/$val/;
                $in_cnt ++;
	        }
            $replace_in = $replace_in . $replace_tmp;
		}
	    while($replace_in =~ /\+<\+\s*([^\+>]*)\s*\+>\+/)		# 	+<+$stg_hash_del_timeout+>+ ==> 10
	    {
            #my $match = $&;
	        my $val = eval($1);
            #print DBG __SUB__ . "REPLACE:" . " $match  $1 => value :  $val  , iterate_cnt : $iterate_cnt\n";
	        $replace_in =~ s/\+<\+\s*([^\+>]*)\s*\+>\+/$val/;
	        $nn_cnt ++;
	    }
    } else {        # this is basic
        while($replace_in =~ /\+<\+\s*([^\+>]*)\s*\+>\+/)		# 	+<+$stg_hash_del_timeout+>+ ==> 10
        {
            #my $match = $&;
                if($num == 369){
	                print DBG __SUB__ . ":" . __LINE__ . " ---- : 1: $1 num $num  $replace_in\n";
	                foreach my $key (sort keys %API) { print DBG __SUB__ . "API1 $key - replace_var_with_value\n"; }
                }
            my $val = eval($1);
            #print DBG __SUB__ . "REPLACE:" . " $match  $1 => value :  $val  , iterate_cnt : $iterate_cnt\n";
            $replace_in =~ s/\+<\+\s*([^\+>]*)\s*\+>\+/$val/;
                if($num == 369){
	                print DBG __SUB__ . ":" . __LINE__ . " ---- : val:$val 1: $1 num $num  $replace_in\n";
	                foreach my $key (sort keys %API) { print DBG __SUB__ . "API1 $key - replace_var_with_value\n"; }
                }
            $in_cnt ++;
        }
    }

    #print DBG __SUB__ . ":" . __LINE__ . " AAAA : before $replace_in\n";
	while($replace_in =~ /(\d+)\s*\+\+\+\+/){		# 	++++     1을 더해 준다. 
		my $temp_num;
		$temp_num = $1;
		$temp_num++;
		$replace_in =~ s/\d+\s*\+\+\+\+/$temp_num/;
		$in_cnt ++;
	}
	#print DBG __SUB__ . ":" . __LINE__ . " ++++ : in_cnt $in_cnt\n";
	while($replace_in =~ /(\d+)\s*\-\-\-\-/){		# 	----     1을 빼준다.
		my $temp_num;
		$temp_num = $1;
		$temp_num--;
		$replace_in =~ s/\d+\s*\-\-\-\-/$temp_num/;
		$in_cnt ++;
	}

	#print DBG __SUB__ . ":" . __LINE__ . " +<+ \$... +>+ : in_cnt $in_cnt\n";
    #print DBG __SUB__ . ":" . __LINE__ . " AAAA : after $replace_in\n";
    if($ln > 10000){
        print "in_cnt matched = $in_cnt $nn_cnt\n";
    }

	return $replace_in;
}

sub change_special_code {
	my ($s) = @_;
	$s =~ s/\{/#\+#\+#\+\+###/g;
	$s =~ s/\}/#\-#\-#\-\-###/g;
	$s =~ s/\\/#\=#\=#\=\=###/g;
	$s =~ s/\n/#\%#\%#\%\%###/g;
	$s =~ s/\"/#\&#\&#\&\&###/g;
	return $s;
}
sub recover_special_code {
	my ($s) = @_;
	$s =~ s/#\+#\+#\+\+###/\{/g;
	$s =~ s/#\-#\-#\-\-###/\}/g;
	$s =~ s/#\=#\=#\=\=###/\\/g;
	$s =~ s/#\%#\%#\%\%###/\n/g;
	$s =~ s/#\&#\&#\&\&###/\"/g;
	return $s;
}
sub traverse_hash_tree_to_recover_special_code {
	my ($TAXA_TREE,$vn,$lstr,$fh)    = @_;
	#print "sub $TAXA_TREE\n";
	foreach my $key (sort{$a<=>$b} keys %{$TAXA_TREE}) {
		if (ref $TAXA_TREE->{$key} eq 'HASH') {
			#print "K:$key lstr=$lstr\n";
			traverse_hash_tree_to_recover_special_code($TAXA_TREE->{$key},$vn,$lstr . "\{" . recover_special_code($key) . "\}",$fh);
		} else {
			#print "$lstr $key = $TAXA_TREE->{$key}\n";
			print $fh "\$$vn$lstr\{" . recover_special_code($key) ."\}=\"" . recover_special_code($TAXA_TREE->{$key}) . "\"\n";
		}
	}
}
sub traverse_hash_tree {
	my ($TAXA_TREE,$vn,$lstr,$fh)    = @_;
	traverse_hash_tree_to_recover_special_code($TAXA_TREE,$vn,$lstr,$fh);
}

#
#======= main MAIN ==============
#

print STDERR "perlscript=$0\n";

our $tttime;
our $cchange_start_time;
our $outputdir = "OUTPUT";
our %local_var_set;
our $optionDebug=0;
our $optionDebugDetail=0;
our $optionPerformance=1;
our $optionNoLog=0;
my $filename="default.GV";
my $stcfilename="default.stc";
our $cga_rdl_version_input_file   = "";
our $cgaRdlVersionMajor;
our $cgaRdlVersionMinor;
our $cgaRdlVersionDev;

GetOptions (
		"inputstci=s"   => \$stcfilename,      # string
		"outputdb=s"   => \$filename,      # string
		"cga_rdl_version_input=s"   => \$cga_rdl_version_input_file,      # string
        "debug" => sub { $optionDebug = 1 },   # flag
        "debugdetail" => sub { $optionDebug = 1;  $optionDebugDetail = 1;  },   # flag
        "original" => sub { $optionPerformance = 0 },   # flag
        "nolog" => sub { $optionNoLog = 1 },   # flag
		"verbose|help"  => sub { $help = 1 })   # flag
or  die(help() . "Error in command line arguments\n");

if($help == 1){
	help();
	exit();
}

if($cga_rdl_version_input_file ne ""){
    my $_dir = dirname($0);
    if(open(my $verfh, "<", "$_dir\/$cga_rdl_version_input_file")){
        my $s = <$verfh>;
        $s =~ /\s*\D*\s*(\d+)\.(\d+)\.(\d+)\s*$/;
        $cgaRdlVersionMajor = $1;
        $cgaRdlVersionMinor = $2;
        $cgaRdlVersionDev = $3;
        print STDERR "CGA_RDL CURRENT VERSION $s => Major $cgaRdlVersionMajor, Minor $cgaRdlVersionMinor, Dev $cgaRdlVersionDev\n";
        close($verfh);
    } else {
        print STDERR "cga rdl version file ($_dir\/$cga_rdl_version_input_file) does not exist.\n";
        $cga_rdl_version_input_file = "";
    }
}


mkdir "OUTPUT";
mkdir "OUTPUT/stc";
mkdir "OUTPUT/tmp";
# set the variables from file
#print "arguments count : $#ARGV\n";
#print STDERR "arguments count : $#ARGV\n";
#($filename,$stcfilename) = (@ARGV);
#if($stcfilename eq ""){
#$filename = "default.GVm";
#$stcfilename = "default.stc";
#}

## init file open
if($optionNoLog){
    open(DBG,">/dev/null");
    open(TIME_DBG,">/dev/null");
    open(VAR_DBG,">/dev/null");
} else {
    open(DBG,">debug.log");
    open(TIME_DBG,">time_debug.log");
    open(VAR_DBG,">var_debug.log");
}


print "fname = $filename , stc fname = $stcfilename\n";
print STDERR "fname = $filename , stc fname = $stcfilename\n";
open(FH, "<",$filename) or die "Can't open < $filename: $!";
my $lcnt = 0 ;
while(<FH>){
	$s = $s_org = $_;
	chop($s);
	eval $s;
	$s =~ /^\$([^\{]+)/;
	$hashName{$1} = "Done";
	if($1 eq "gCan"){ print "== $s\n"; }
	#recover_hash_value(\%{$vname},$s);
	#LOG1 print $s_org;
	#if($lcnt >50){ last; }
	#$lcnt++;
}
close(FH);


if($cga_rdl_version_input_file ne ""){
    my $tmpMajor = 0;
    my $tmpMinor = 0;
    my $tmpDev = 0;
    my $s = $gTitle{"VARIABLE"}{"CGA_RDL_Version"};
    $s =~ /\s*\D*\s*(\d+)\.(\d+)\.(\d+)\s*$/;
    $tmpMajor = $1;
    $tmpMinor = $2;
    $tmpDev = $3;
    print STDERR "cga_rdl FILE VERSION $s => Major $tmpMajor, Minor $tmpMinor, Dev $tmpDev\n";
    if($cgaRdlVersionMajor != $tmpMajor){ versionMismatch(); exit(4); }
    if($cgaRdlVersionMinor != $tmpMinor){ versionMismatch(); exit(4); }
    if($cgaRdlVersionDev != $tmpDev){ versionMismatch(); exit(4); }
}


foreach my $key (keys %gVariables){
	$$key = $gVariables{$key};
}



foreach my $key (keys %hashName){
	print VAR_DBG "----[$key]----\n";
	traverse_hash_tree(\%{$key},$key,"",VAR_DBG);
}

if($stcfilename =~ /\.stc$/){
	CChange($stcfilename,"stc","","DEBUG_ON");
} elsif($stcfilename =~ /\.stcI$/){
	IChange($stcfilename,"stc");
}

exit;

print \%{gCol} . "\n";
print getHashRef("gCol") . "\n";
foreach my $key (sort{$a<=>$b} keys %{getHashRef("gCol")}) { print "$key  ";} print "\n";
print $gCol{9} . "\n";
print \%{"gCol{9}"} . "\n";
print getHashRef("gCol{9}") . "\n";
$tmp = $gCol{9};
foreach my $key (sort{$a<=>$b} keys %$tmp) { print "A$key  ";} print "\n";
foreach my $key (sort{$a<=>$b} keys %{(getHashRef("gCol"))->{9}}) { print "B$key  ";} print "\n";
foreach my $key (sort{$a<=>$b} keys %{getHashRef("gCol{9}")}) { print "C$key  ";} print "\n";
foreach my $key (sort{$a<=>$b} keys %{getHashRef("gCan{3-4}")}) { print "C$key  ";} print "\n";
foreach my $key (sort{$a<=>$b} keys %{getHashRef("gCan{9}")}) { print "C$key  ";} print "\n";
foreach my $key (sort{$a<=>$b} keys %{getHashRef("gCan{1}")}) { print "C$key  ";} print "\n";
foreach my $key (sort{$a<=>$b} keys %{getHashRef("gCan{2}")}) { print "C$key  ";} print "\n";
getHashRef(gCan)->{100}->{10} = 11111;
foreach my $key (sort{$a<=>$b} keys %{getHashRef("gCan")}) { print "C$key  ";} print "\n";
foreach my $key (sort{$a<=>$b} keys %{getHashRef("gCan{100}")}) { print "C$key  ";} print "\n";

my $comment =  <<END_COMMENT;
		foreach my $key (keys %hashName){
			print VAR_DBG "----[$key]----\n";
			traverse_hash_tree(\%{$key},$key,"",VAR_DBG);
		}
END_COMMENT

