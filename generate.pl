#!c:/perl/bin -w

#################################################
#auther:jackylee
#date:2013-4-10
#function:auto fill summiting documents
#################################################

use strict;
#use diagnostics;

use Cwd;
use File::Copy;
use File::Find;
use File::Spec;
use File::Basename;
use File::Path qw(make_path remove_tree);
use WIN32::OLE;
use Win32::OLE::Const 'Microsoft Word';
use Win32::OLE::Const 'Microsoft Excel';
use Win32::File;


my $rootDir = getcwd;
my $srcDir = "打包模板";
my $destDir = "提交文件";
my ($ab_flag, $c_flag) = ('需求说明书', '需求申请单');



################################################################################################################
#step1: judge ab or c  level
print "\n";
opendir DIR, $rootDir or die("can't open $rootDir. ERR:$!");
my @file = readdir DIR;
closedir DIR;

my @file1 = grep /\.doc/, @file;

my @tmp_ab = grep /$ab_flag/, @file1;
my @tmp_c  = grep /$c_flag/, @file1;

my $tmp_ab = @tmp_ab;
my $tmp_c  = @tmp_c; 


if ($tmp_ab && $tmp_c){
		print "当前目录$rootDir 存在两个需求说明书,请选择:\n", 
		 "(1): ", $tmp_ab[0],"\n",
		 "(2): ", $tmp_c[0],"\n";
		 
	while(1){	
				print "请输入序号:";
				my $num = <STDIN>;
				
				if($num == 1){
					print "选择序号:$num\t$tmp_ab[0]\n";
					$tmp_c = 0;
					last;
				}
				elsif($num == 2){
					print "选择序号:$num\t$tmp_c[0]\n";
					$tmp_ab = 0;
					last;
				}
				else{
					print "序号不存在,";
				}
	}
}
elsif($tmp_ab){
		print "当前目录$rootDir 存在需求说明书:\n", 
					$tmp_ab[0],"\n";
}
elsif($tmp_c){
		print "当前目录$rootDir 存在申请单:\n", 
					$tmp_c[0],"\n";
}
else{
	print "当前目录$rootDir 下不存在\"业务需求说明书\" 或者 \"业务需求申请单\"";
	exit 0;
}


print "\n\n程序正在处理文件，请等待......\n";


################################################################################################################
#step2: delete submit directory
remove_tree("$rootDir/$destDir", {error  => \my $err});
if (@$err) {    
	 for my $diag (@$err) {      
	  	 my ($file, $message) = %$diag;        
	  	 if ($file eq '') {          
	  	      print "general error: $message\n";        
	  	 }        
	  	 else {   
	  	      print "problem unlinking $file: $message\n";     
	  	 }    
	  	} 
}
################################################################################################################  
#step3: create submit directory
make_path("$rootDir/$destDir");

################################################################################################################
#step4: copy files 
if($tmp_ab){
	find(\ &want_ab, "$rootDir/$srcDir/（CEBZH）A、B类需求模板");
	
}
elsif($tmp_c){
	find(\ &want_c, "$rootDir/$srcDir/（CEBZH）C类需求模板");	
	
}

sub want_ab{
		/\.(doc|xls)$/  && copy("$File::Find::name","${rootDir}/${destDir}/")  ;		
}

sub want_c{
		/\.(doc|xls)$/  && copy("$File::Find::name","${rootDir}/${destDir}/")  ;	
}

################################################################################################################
#step5: parse require
my ($sys_name, $req_no , $req_sum, , $req_date , $req_man, $req_tel, $req_depart, $suggest_date);
my ($job_rule, $job_moduel, $job_tuneFace, $job_tuneRelOS, $job_affect, $job_other);

my $word = Win32::OLE->GetActiveObject('Word.Application') || Win32::OLE->new('Word.Application', 'Quit'); #open word document
my $document;
        
$word->{Visible} = 0;


        
$sys_name = "cebzh";
$req_no = "xxxxxx";
$req_sum = "default";     
my $req_tmp;   
 
#print "$tmp_ab";        
if($tmp_ab){
	
	#print "${rootDir}/$tmp_ab[0]";
	$document= $word->Documents->Open("${rootDir}/$tmp_ab[0]") 
							|| die("Unable to open document ${rootDir}/$tmp_ab[0]", Win32::OLE->LastError());
							
	my $table_ab = $document->Tables(1);
	
	#get sys_name
	$req_tmp = $table_ab->Cell(1,2)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$sys_name = "$1";
	
	#get req_depart
	$req_tmp = $table_ab->Cell(1,4)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$req_depart = "$1";
	
	#get req_no
	$req_tmp = $table_ab->Cell(2,2)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$req_no = "$1";

	#get req_sum
	$req_tmp = $table_ab->Cell(2,4)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$req_sum = "$1";
	
	#get req_man
	$req_tmp = $table_ab->Cell(3,2)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$req_man = "$1";
	
	#get req_tel
	$req_tmp = $table_ab->Cell(3,4)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$req_tel = "$1";				
	
	
	#get req_date
	$req_tmp = $table_ab->Cell(4,2)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$req_date = "$1";
	
	#get suggest_date
	$req_tmp = $table_ab->Cell(4,4)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$suggest_date = "$1";			
	
	
	$document->Close;
	
	#print "\n",
	#			"系统名称:",$sys_name,"\n",
	#			"需求编号:",$req_no,"\n",
	#			"需求名称:",$req_sum,"\n";							
}


my $table_c;
if($tmp_c){
	$document= $word->Documents->Open("${rootDir}/$tmp_c[0]") 
							|| die("Unable to open document ${rootDir}/$tmp_c[0] ", Win32::OLE->LastError());	
	
	$table_c= $document->Tables(1);
	
	#get sys_name
	$req_tmp = $table_c->Cell(1,2)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$sys_name = "$1";
	
	#get req_depart
	$req_tmp = $table_c->Cell(1,4)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$req_depart = "$1";
	
	#get req_no
	$req_tmp = $table_c->Cell(2,2)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$req_no = "$1";

	#get req_sum
	$req_tmp = $table_c->Cell(2,4)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$req_sum = "$1";
	
	#get req_man
	$req_tmp = $table_c->Cell(3,2)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$req_man = "$1";
	
	#get req_tel
	$req_tmp = $table_c->Cell(3,4)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$req_tel = "$1";				
	
	
	#get req_date
	$req_tmp = $table_c->Cell(4,2)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$req_date = "$1";
	
	#get suggest_date
	$req_tmp = $table_c->Cell(4,4)->Range->{Text};
	$req_tmp =~ /^(.+)\s(.+)$/;
	$suggest_date = "$1";			
	
	
	$document->Close;
	
	#print "\n",
	#			"系统名称:",$sys_name,"\n",
	#			"需求编号:",$req_no,"\n",
	#			"需求名称:",$req_sum,"\n";
	

	
}


################################################################################################################
#step6: rename files  ,then set writeable 
my @filename;
my ($fpath, $fname, $suffix);
my @key_name ;


print "\n";
my $destfulldir = "$rootDir/$destDir";
opendir DIR,  "$destfulldir" or die("can't open $destfulldir. ERR:$!");
my @destfile = readdir DIR;
closedir DIR;


if($tmp_ab){

		@key_name = ('详细设计说明书',
								 '应用系统安装投产审批表',
								 '系统投产变更实施控制表',
								 '程序修改清单',
								 '业务需求说明书',
								 '业务需求分析说明书',
								 '用户验收测试报告',
								 '技术测试报告',
								 '代码走查单',
								 '设计评审检查单',
								 '投产评审检查单',
								 '需求评审检查单',
								 '软件变更',
								 '设计评审报告',
								 '需求评审报告'								 
								 );
			
			foreach my $s1 (@key_name){
				
				@filename = grep /$s1/, @destfile;
				($fpath, $fname, $suffix) = fileparse("$destfulldir/$filename[0]", qr/\.[^.]*/);
				move("$destfulldir/$filename[0]", "${destfulldir}/${sys_name}_${s1}_${req_no}${req_sum}$suffix");	
				Win32::File::SetAttributes("${destfulldir}/${sys_name}_${s1}_${req_no}${req_sum}$suffix",NORMAL)
																		 or die "Can't set attributes for ${destfulldir}/${sys_name}_${s1}_${req_no}${req_sum}$suffix.";
			
		}
		
}


if($tmp_c){

				@key_name = ('业务需求申请单',
								 '用户验收测试报告',
								 '技术测试报告',
								 '技术开发方案',
								 '设计评审报告',
								 '设计评审检查单',
								 '投产评审检查单',
								 '系统投产变更实施控制表',
								 '需求评审报告',
								 '需求评审检查单',
								 '软件变更',
								 '应用系统安装投产审批表'								 
								 );
	
		foreach my $s2 (@key_name){
				
				@filename = grep /$s2/, @destfile;
				($fpath, $fname, $suffix) = fileparse("$destfulldir/$filename[0]", qr/\.[^.]*/);
				move("$destfulldir/$filename[0]", "${destfulldir}/${sys_name}_${s2}_${req_no}${req_sum}$suffix");	
				Win32::File::SetAttributes("${destfulldir}/${sys_name}_${s2}_${req_no}${req_sum}$suffix",NORMAL)
																		 or die "Can't set attributes for ${destfulldir}/${sys_name}_${s2}_${req_no}${req_sum}$suffix.";
						
		}
		
}

################################################################################################################
#step7: get other useful information
use Sys::Hostname;   
  
  my $host = hostname;
  $host =~ /^(.+)-(.+)$/;
  $host = $2;

use POSIX qw(strftime);  
  my $now_string = strftime "%Y-%m-%d", localtime;
  
  
 #print $host;
 #print $now_string;



################################################################################################################
#step8: set doc or xls  content; 
my $excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');#open excel document    
my $book;
my $table;
my $sheet;
$excel->{Visible} = 0;

print "\n";
opendir DIR,  "$destfulldir" or die("can't open $destfulldir. ERR:$!");
@destfile = readdir DIR;
closedir DIR;



my $ss2;

if($tmp_ab){
	
			#详细设计说明书
			$ss2 = $key_name[0];   #doc
			@filename = grep /$ss2/, @destfile;
			$document = $word->Documents->Open("$destfulldir/$filename[0]")
   								 || die("Unable to open document ", Win32::OLE->LastError());   								 
			
			$document->Save;
			$document->Close;		
	
			#应用系统安装投产审批表
			$ss2 = $key_name[1]	;			#doc
			@filename = grep /$ss2/, @destfile;
			$document = $word->Documents->Open("$destfulldir/$filename[0]")
   								 || die("Unable to open document ", Win32::OLE->LastError());

			$table = $document->Tables(1);
			$table->Cell(1,2)->Range->{Text} = $sys_name;
			
			$document->Save;			
			$document->Close;		

			#系统投产变更实施控制表    #xls
			$ss2 = $key_name[2]  ;   
			@filename = grep /$ss2/, @destfile;
			$book = $excel->Workbooks->Open("$destfulldir/$filename[0]") 
						|| die("Unable to open document ", Win32::OLE->LastError());
			$book->{CheckCompatibility } = 0; #turn off check compatibility			
						
			$book->Save	;				
			$book->Close;									 								 
			
			
			#程序修改清单	
			$ss2 = $key_name[3]  ;     #doc
			@filename = grep /$ss2/, @destfile;
			$document = $word->Documents->Open("$destfulldir/$filename[0]")
   								 || die("Unable to open document ", Win32::OLE->LastError());
   					 
			$table = $document->Tables(1);
			$table->Cell(1,2)->Range->{Text} = $sys_name;
			$table->Cell(2,2)->Range->{Text} = $req_no;
			$table->Cell(2,4)->Range->{Text} = $req_sum;
			$table->Cell(3,2)->Range->{Text} = $host;
			
			$table->Cell(1,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(1,2)->Range->Font->{Name} = "宋体";
			
			$table->Cell(2,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(2,2)->Range->Font->{Name} = "宋体";
			
			$table->Cell(2,4)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(2,4)->Range->Font->{Name} = "宋体";
			
			$table->Cell(3,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(3,2)->Range->Font->{Name} = "宋体";
				
			$document->Save;		
			$document->Close;							

			#业务需求说明书        #doc
			$ss2 = $key_name[4];
			@filename = grep /$ss2/, @destfile;
			$document = $word->Documents->Open("$destfulldir/$filename[0]")
   								 || die("Unable to open document ", Win32::OLE->LastError());
			
			$document->Save;
			$document->Close;		


			#业务需求分析说明书    #doc
			$ss2 = $key_name[5];
			@filename = grep /$ss2/, @destfile;
			$document = $word->Documents->Open("$destfulldir/$filename[0]")
   								 || die("Unable to open document ", Win32::OLE->LastError());
			
			$document->Save;
			$document->Close;		


			#用户验收测试报告      #doc
			$ss2 = $key_name[6];
			@filename = grep /$ss2/, @destfile;
			$document = $word->Documents->Open("$destfulldir/$filename[0]")
   								 || die("Unable to open document ", Win32::OLE->LastError());
			
			$document->Save;
			$document->Close;		
			
			
			#技术测试报告    #doc
			$ss2 = $key_name[7];
			@filename = grep /$ss2/, @destfile;
			$document = $word->Documents->Open("$destfulldir/$filename[0]")
   								 || die("Unable to open document ", Win32::OLE->LastError());
   								 								 
			$table = $document->Tables(1);
			$table->Cell(1,2)->Range->{Text} = $sys_name;
			$table->Cell(1,4)->Range->{Text} = "${req_no}/${req_sum}";
			$table->Cell(2,2)->Range->{Text} = $host;
			$table->Cell(2,4)->Range->{Text} = $now_string;
			$table->Cell(2,6)->Range->{Text} = $host;
			
			$table->Cell(1,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(1,2)->Range->Font->{Name} = "宋体";
			$table->Cell(1,2)->Range->ParagraphFormat->{Alignment} = wdAlignParagraphCenter ;
			
			$table->Cell(1,4)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(1,4)->Range->Font->{Name} = "宋体";
			$table->Cell(1,4)->Range->ParagraphFormat->{Alignment} = wdAlignParagraphCenter ;
			
			$table->Cell(2,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(2,2)->Range->Font->{Name} = "宋体";
			$table->Cell(2,2)->Range->ParagraphFormat->{Alignment} = wdAlignParagraphCenter ;
			
			$table->Cell(2,4)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(2,4)->Range->Font->{Name} = "宋体";
			$table->Cell(2,4)->Range->ParagraphFormat->{Alignment} = wdAlignParagraphCenter ;			
			
			$table->Cell(2,6)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(2,6)->Range->Font->{Name} = "宋体";
			$table->Cell(2,6)->Range->ParagraphFormat->{Alignment} = wdAlignParagraphCenter ;			
			
			$document->Save;						
			$document->Close;			


			#代码走查单    #xls
			$ss2 = $key_name[8];
			@filename = grep /$ss2/, @destfile;
			$book = $excel->Workbooks->Open("$destfulldir/$filename[0]") 
						|| die("Unable to open document ", Win32::OLE->LastError());
			$book->{CheckCompatibility } = 0; #turn off check compatibility			
						
			$sheet = $book->Worksheets("代码走查");
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 

			$sheet->Range("D3")->{Value} = $req_sum;
			$sheet->Range("F3")->{Value} = $req_no;
			$sheet->Range("I3")->{Value} = "1.0";
			$sheet->Range("D4")->{Value} = $req_date;
			$sheet->Range("D4")->{Value} = $host;		
			
			$book->Save	;				
			$book->Close;	


			#设计评审检查单       #xls
			$ss2 = $key_name[9];
			@filename = grep /$ss2/, @destfile;
			$book = $excel->Workbooks->Open("$destfulldir/$filename[0]") 
						|| die("Unable to open document ", Win32::OLE->LastError());
			$book->{CheckCompatibility } = 0; #turn off check compatibility			

			$sheet = $book->Worksheets("开发设计阶段评审检查单");
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 
			
			$sheet->Range("B2")->{Value} = $req_no;
			$sheet->Range("D2")->{Value} = "项目名称: ${req_sum}";
			$sheet->Range("F2")->{Value} = "陈明垓";
			$sheet->Range("B3")->{Value} = $host;
			$sheet->Range("D3")->{Value} = "检查日期: ${req_date}";
			$sheet->Range("F3")->{Value} = "0.5小时";		
			$sheet->Range("A54")->{Value} = "总体情况良好";
			$sheet->Range("A56")->{Value} = "无记录";
			$sheet->Range("A58")->{Value} = "无";			
		

			$sheet = $book->Worksheets("运行设计阶段评审检查单");
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 
			
			$sheet->Range("B2")->{Value} = $req_no;
			$sheet->Range("D2")->{Value} = "项目名称: ${req_sum}";
			$sheet->Range("F2")->{Value} = "陈明垓";
			$sheet->Range("B3")->{Value} = $host;
			$sheet->Range("D3")->{Value} = "检查日期: ${req_date}";
			$sheet->Range("F3")->{Value} = "0.5小时";		
			$sheet->Range("A40")->{Value} = "总体情况良好";
			$sheet->Range("A42")->{Value} = "无记录";
			$sheet->Range("A44")->{Value} = "无";			
	
				
			$sheet = $book->Worksheets("安全设计阶段评审检查表");
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 
			
			$sheet->Range("B2")->{Value} = $req_no;
			$sheet->Range("D2")->{Value} = "项目名称: ${req_sum}";
			$sheet->Range("F2")->{Value} = "陈明垓";
			$sheet->Range("B3")->{Value} = $host;
			$sheet->Range("D3")->{Value} = "检查日期: ${req_date}";
			$sheet->Range("F3")->{Value} = "0.5小时";		
			$sheet->Range("A34")->{Value} = "总体情况良好";
			$sheet->Range("A36")->{Value} = "无记录";
			$sheet->Range("A38")->{Value} = "无";	
			
			$book->Save	;								
			$book->Close;	
			
			
			#投产评审检查单           #xls
			$ss2 = $key_name[10] ;      
			@filename = grep /$ss2/, @destfile;
			$book = $excel->Workbooks->Open("$destfulldir/$filename[0]") 
						|| die("Unable to open document ", Win32::OLE->LastError());
   								 
			$book->{CheckCompatibility } = 0; #turn off check compatibility			

			$sheet = $book->Worksheets("投产评审检查单");
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 
			
			$sheet->Range("C2")->{Value} = $req_no;
			$sheet->Range("E2")->{Value} = $req_sum;
			$sheet->Range("G2")->{Value} = "陈明垓";
			$sheet->Range("C3")->{Value} = $host;
			$sheet->Range("E3")->{Value} = $req_date;
			$sheet->Range("G3")->{Value} = "0.5小时";

			$sheet->Range("B41")->{Value} = "通过";
			$sheet->Range("B43")->{Value} = "无";
			$sheet->Range("B45")->{Value} = "无";
			
			$book->Save	;									
			$book->Close;			
			

			#需求评审检查单      #xls
			$ss2 = $key_name[11];  
			@filename = grep /$ss2/, @destfile;
			$book = $excel->Workbooks->Open("$destfulldir/$filename[0]") 
						|| die("Unable to open document ", Win32::OLE->LastError());
			$book->{CheckCompatibility } = 0; #turn off check compatibility			

			$sheet = $book->Worksheets("需求评审检查单");
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 

			
			$sheet->Range("B4")->{Value} = $req_no;
			$sheet->Range("D4")->{Value} = $req_sum;
			$sheet->Range("G4")->{Value} = "陈明垓";
			$sheet->Range("B5")->{Value} = $host;
			$sheet->Range("D5")->{Value} = $req_date;
			$sheet->Range("G5")->{Value} = "0.5小时";
			
			$sheet->Range("A64")->{Value} = "总体情况较好";
			$sheet->Range("A66")->{Value} = "无";
			$sheet->Range("A68")->{Value} = "无";
			
			$book->Save	;				
			$book->Close;			
			
			
			
			#软件变更           #doc
			$ss2 = $key_name[12] ;       #doc
			@filename = grep /$ss2/, @destfile;
			$document = $word->Documents->Open("$destfulldir/$filename[0]")
   								 || die("Unable to open document ", Win32::OLE->LastError());
   								 
			$table = $document->Tables(1);				
			$table->Cell(1,2)->Range->{Text} = $sys_name;
			$table->Cell(2,2)->Range->{Text} = $host;
			$table->Cell(2,4)->Range->{Text} = $now_string;
	
		
			$document->Paragraphs->Add( {Range => $document->Paragraphs(1)->Range}) ; 
			$document->Paragraphs(1)->Range->{Text} = "编号: ${req_no}_${req_sum}_${now_string}\n"; 
			$document->Paragraphs(1)->Range->Font->{Italic } = 0 ;
			$document->Paragraphs(2)->Range->{Text} = "";

			$table->Cell(1,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(1,2)->Range->Font->{Name} = "宋体";
			$table->Cell(1,2)->Range->Font->{Italic } = 0 ;
			

			$table->Cell(2,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(2,2)->Range->Font->{Name} = "宋体";
			
			$table->Cell(2,4)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(2,4)->Range->Font->{Name} = "宋体";
			$table->Cell(2,4)->Range->Font->{Italic } = 0 ;	
			
			$document->Save;				
			$document->Close;		
			
			#设计评审报告       #xls
			$ss2 = $key_name[13];
			@filename = grep /$ss2/, @destfile;
			$book = $excel->Workbooks->Open("$destfulldir/$filename[0]") 
						|| die("Unable to open document ", Win32::OLE->LastError());
			
			
			$book->{CheckCompatibility } = 0; #turn off check compatibility			
			$sheet = $book->Worksheets(2);
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 
			
			$sheet->Range("E4")->{Value} = "${req_no}_${req_sum}";
			$sheet->Range("D34")->{Value} = "通过";
			$sheet->Range("E41")->{Value} = "无";
			$sheet->Range("K41")->{Value} = "";
			
			$sheet->Range("E4")->Font->{Color} = xlThemeColorDark1;
			$sheet->Range("D34")->Font->{Italic} = 0 ;
			$sheet->Range("E41")->Font->{Color} = xlThemeColorDark1;	
			$sheet->Range("E41")->Font->{Bold} = 0;
			
			
			$book->Save	;			
			$book->Close;		
			
						
			#需求评审报告		    #xls
			$ss2 = $key_name[14];
			@filename = grep /$ss2/, @destfile;
			$book = $excel->Workbooks->Open("$destfulldir/$filename[0]") 
						|| die("Unable to open document ", Win32::OLE->LastError());
			$book->{CheckCompatibility } = 0; #turn off check compatibility			
						
			$sheet = $book->Worksheets("技术评审报告");
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 

			$sheet->Range("E4")->{Value} = "${req_no}_${req_sum}";
			$sheet->Range("D34")->{Value} = "通过";
			$sheet->Range("E41")->{Value} = "无";
			$sheet->Range("K41")->{Value} = "";
			
			$sheet->Range("E4")->Font->{Color} = xlThemeColorDark1;
			$sheet->Range("D34")->Font->{Italic} = 0 ;
			$sheet->Range("E41")->Font->{Color} = xlThemeColorDark1;	
			$sheet->Range("E41")->Font->{Bold} = 0;
			
			
			$book->Save	;					
			$book->Close;	
					
}


if($tmp_c){
	
	
			#业务需求申请单
			$ss2 = $key_name[0];   #doc
			@filename = grep /$ss2/, @destfile;
			$document = $word->Documents->Open("$destfulldir/$filename[0]")
   								 || die("Unable to open document ", Win32::OLE->LastError());
			
			$document->Save;						
			$document->Close;		
	
			#用户验收测试报告
			$ss2 = $key_name[1]	;			#doc
			@filename = grep /$ss2/, @destfile;
			$document = $word->Documents->Open("$destfulldir/$filename[0]")
   								 || die("Unable to open document ", Win32::OLE->LastError());
			
			$document->Save;							
			$document->Close;		

			#技术测试报告
			$ss2 = $key_name[2]  ;   #doc
			@filename = grep /$ss2/, @destfile;
			$document = $word->Documents->Open("$destfulldir/$filename[0]")
   								 || die("Unable to open document ", Win32::OLE->LastError());
   								 								 
			$table = $document->Tables(1);
			$table->Cell(1,2)->Range->{Text} = $sys_name;
			$table->Cell(1,4)->Range->{Text} = "${req_no}/${req_sum}";
			$table->Cell(2,2)->Range->{Text} = $host;
			$table->Cell(2,4)->Range->{Text} = $now_string;
			$table->Cell(2,6)->Range->{Text} = $host;
			
			
			#use Win32::OLE::Const 'Microsoft Word';
			$table->Cell(1,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(1,2)->Range->Font->{Name} = "宋体";
			$table->Cell(1,2)->Range->ParagraphFormat->{Alignment} = wdAlignParagraphCenter ;
			
			$table->Cell(1,4)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(1,4)->Range->Font->{Name} = "宋体";
			$table->Cell(1,4)->Range->ParagraphFormat->{Alignment} = wdAlignParagraphCenter ;
			
			$table->Cell(2,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(2,2)->Range->Font->{Name} = "宋体";
			$table->Cell(2,2)->Range->ParagraphFormat->{Alignment} = wdAlignParagraphCenter ;
			
			$table->Cell(2,4)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(2,4)->Range->Font->{Name} = "宋体";
			$table->Cell(2,4)->Range->ParagraphFormat->{Alignment} = wdAlignParagraphCenter ;			
			
			$table->Cell(2,6)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(2,6)->Range->Font->{Name} = "宋体";
			$table->Cell(2,6)->Range->ParagraphFormat->{Alignment} = wdAlignParagraphCenter ;			
			
			$document->Save;							
			$document->Close;	
			
			
			#技术开发方案	
			$ss2 = $key_name[3]  ;     #doc
			@filename = grep /$ss2/, @destfile;
			$document = $word->Documents->Open("$destfulldir/$filename[0]")
   								 || die("Unable to open document ", Win32::OLE->LastError());
   					 
			$table = $document->Tables(1);
			$table->Cell(1,2)->Range->{Text} = $sys_name;
			$table->Cell(2,2)->Range->{Text} = $req_no;
			$table->Cell(2,4)->Range->{Text} = $req_sum;
			$table->Cell(5,2)->Range->{Text} = $host;
			$table->Cell(5,4)->Range->{Text} = $now_string;
			$table->Cell(7,2)->Range->{Text} = $host;
			
			
			#use Win32::OLE::Const 'Microsoft Word';
			$table->Cell(1,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(1,2)->Range->Font->{Name} = "宋体";
			
			$table->Cell(2,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(2,2)->Range->Font->{Name} = "宋体";
			
			$table->Cell(2,4)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(2,4)->Range->Font->{Name} = "宋体";
			
			$table->Cell(5,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(5,2)->Range->Font->{Name} = "宋体";
			
			$table->Cell(5,4)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(5,4)->Range->Font->{Name} = "宋体";
			
			$table->Cell(7,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(7,2)->Range->Font->{Name} = "宋体";			
			
			$document->Save;			
			$document->Close;							

			#设计评审报告
			$ss2 = $key_name[4];
			@filename = grep /$ss2/, @destfile;
			$book = $excel->Workbooks->Open("$destfulldir/$filename[0]") 
						|| die("Unable to open document ", Win32::OLE->LastError());
			
			
			$book->{CheckCompatibility } = 0; #turn off check compatibility			
			$sheet = $book->Worksheets(2);
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 
			
			$sheet->Range("E4")->{Value} = "${req_no}_${req_sum}";
			$sheet->Range("D46")->{Value} = "通过";
			$sheet->Range("E53")->{Value} = "无";
			$sheet->Range("K53")->{Value} = "";
			
			$sheet->Range("E4")->Font->{Color} = xlThemeColorDark1;
			$sheet->Range("E53")->Font->{Italic} = 0 ;
			$sheet->Range("E53")->Font->{Color} = xlThemeColorDark1;	
			$sheet->Range("E53")->Font->{Bold} = 0;
			
			
			$book->Save	;			
			$book->Close;		


			#设计评审检查单
			$ss2 = $key_name[5];
			@filename = grep /$ss2/, @destfile;
			$book = $excel->Workbooks->Open("$destfulldir/$filename[0]") 
						|| die("Unable to open document ", Win32::OLE->LastError());
			$book->{CheckCompatibility } = 0; #turn off check compatibility			

			$sheet = $book->Worksheets("开发设计阶段评审检查单");
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 
			
			$sheet->Range("B2")->{Value} = $req_no;
			$sheet->Range("D2")->{Value} = "项目名称: ${req_sum}";
			$sheet->Range("F2")->{Value} = "陈明垓";
			$sheet->Range("B3")->{Value} = $host;
			$sheet->Range("D3")->{Value} = "检查日期: ${req_date}";
			$sheet->Range("F3")->{Value} = "0.5小时";		
			$sheet->Range("A53")->{Value} = "总体情况良好";
			$sheet->Range("A55")->{Value} = "无记录";
			$sheet->Range("A57")->{Value} = "无";			
		

			$sheet = $book->Worksheets("运行设计阶段评审检查单");
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 
			
			$sheet->Range("B2")->{Value} = $req_no;
			$sheet->Range("D2")->{Value} = "项目名称: ${req_sum}";
			$sheet->Range("F2")->{Value} = "陈明垓";
			$sheet->Range("B3")->{Value} = $host;
			$sheet->Range("D3")->{Value} = "检查日期: ${req_date}";
			$sheet->Range("F3")->{Value} = "0.5小时";		
			$sheet->Range("A40")->{Value} = "总体情况良好";
			$sheet->Range("A42")->{Value} = "无记录";
			$sheet->Range("A44")->{Value} = "无";			
	
				
			$sheet = $book->Worksheets("安全设计阶段评审检查表");
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 
			
			$sheet->Range("B2")->{Value} = $req_no;
			$sheet->Range("D2")->{Value} = "项目名称: ${req_sum}";
			$sheet->Range("F2")->{Value} = "陈明垓";
			$sheet->Range("B3")->{Value} = $host;
			$sheet->Range("D3")->{Value} = "检查日期: ${req_date}";
			$sheet->Range("F3")->{Value} = "0.5小时";		
			$sheet->Range("A34")->{Value} = "总体情况良好";
			$sheet->Range("A36")->{Value} = "无记录";
			$sheet->Range("A38")->{Value} = "无";	
			
			$book->Save	;								
			$book->Close;	


			#投产评审检查单
			$ss2 = $key_name[6];
			@filename = grep /$ss2/, @destfile;
			$book = $excel->Workbooks->Open("$destfulldir/$filename[0]") 
						|| die("Unable to open document ", Win32::OLE->LastError());
			$book->{CheckCompatibility } = 0; #turn off check compatibility			

			$sheet = $book->Worksheets("投产评审检查单");
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 
			
			$sheet->Range("C2")->{Value} = $req_no;
			$sheet->Range("E2")->{Value} = $req_sum;
			$sheet->Range("G2")->{Value} = "陈明垓";
			$sheet->Range("C3")->{Value} = $host;
			$sheet->Range("E3")->{Value} = $req_date;
			$sheet->Range("G3")->{Value} = "0.5小时";

			$sheet->Range("B41")->{Value} = "通过";
			$sheet->Range("B43")->{Value} = "无";
			$sheet->Range("B45")->{Value} = "无";
			
			$book->Save	;									
			$book->Close;			
			
			
			#系统投产变更实施控制表
			$ss2 = $key_name[7];
			@filename = grep /$ss2/, @destfile;
			$book = $excel->Workbooks->Open("$destfulldir/$filename[0]") 
						|| die("Unable to open document ", Win32::OLE->LastError());
			$book->{CheckCompatibility } = 0; #turn off check compatibility			

			$book->Save	;							
			$book->Close;		


			#需求评审报告
			$ss2 = $key_name[8];
			@filename = grep /$ss2/, @destfile;
			$book = $excel->Workbooks->Open("$destfulldir/$filename[0]") 
						|| die("Unable to open document ", Win32::OLE->LastError());
			$book->{CheckCompatibility } = 0; #turn off check compatibility			
						
			$sheet = $book->Worksheets("技术评审报告");
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 

			$sheet->Range("E4")->{Value} = "${req_no}_${req_sum}";
			$sheet->Range("D46")->{Value} = "通过";
			$sheet->Range("E53")->{Value} = "无";
			$sheet->Range("K53")->{Value} = "";
			
			$sheet->Range("E4")->Font->{Color} = xlThemeColorDark1;
			$sheet->Range("E53")->Font->{Italic} = 0 ;
			$sheet->Range("E53")->Font->{Color} = xlThemeColorDark1;	
			$sheet->Range("E53")->Font->{Bold} = 0;
			
			
			$book->Save	;			
			$book->Close;	

			#需求评审检查单
			$ss2 = $key_name[9];
			@filename = grep /$ss2/, @destfile;
			$book = $excel->Workbooks->Open("$destfulldir/$filename[0]") 
						|| die("Unable to open document ", Win32::OLE->LastError());
			$book->{CheckCompatibility } = 0; #turn off check compatibility			

			$sheet = $book->Worksheets("需求评审检查单");
			$sheet->Cells->{NumberFormatLocal} = "@";  # set range format as strings 

			
			$sheet->Range("B4")->{Value} = $req_no;
			$sheet->Range("D4")->{Value} = $req_sum;
			$sheet->Range("G4")->{Value} = "陈明垓";
			$sheet->Range("B5")->{Value} = $host;
			$sheet->Range("D5")->{Value} = $req_date;
			$sheet->Range("G5")->{Value} = "0.5小时";
			
			$sheet->Range("A64")->{Value} = "总体情况较好";
			$sheet->Range("A66")->{Value} = "无";
			$sheet->Range("A68")->{Value} = "无";
			
			$book->Save	;				
			$book->Close;			
			
			
			#软件变更
			$ss2 = $key_name[10] ;       #doc
			@filename = grep /$ss2/, @destfile;
			$document = $word->Documents->Open("$destfulldir/$filename[0]")
   								 || die("Unable to open document ", Win32::OLE->LastError());
   								 
			$table = $document->Tables(1);				
			$table->Cell(1,2)->Range->{Text} = $sys_name;
			$table->Cell(2,2)->Range->{Text} = $host;
			$table->Cell(2,4)->Range->{Text} = $now_string;
	
		
			$document->Paragraphs->Add( {Range => $document->Paragraphs(1)->Range}) ; 
			$document->Paragraphs(1)->Range->{Text} = "编号: ${req_no}_${req_sum}_${now_string}\n"; 
			$document->Paragraphs(1)->Range->Font->{Italic } = 0 ;
			$document->Paragraphs(2)->Range->{Text} = "";

			#use Win32::OLE::Const 'Microsoft Word';
			$table->Cell(1,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(1,2)->Range->Font->{Name} = "宋体";
			$table->Cell(1,2)->Range->Font->{Italic } = 0 ;
			


			$table->Cell(2,2)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(2,2)->Range->Font->{Name} = "宋体";
			
			$table->Cell(2,4)->Range->Font->{Color} = wdColorBlack;
			$table->Cell(2,4)->Range->Font->{Name} = "宋体";
			$table->Cell(2,4)->Range->Font->{Italic } = 0 ;	
			
			$document->Save;				
			$document->Close;		

			#应用系统安装投产审批表
			$ss2 = $key_name[11];  #doc
			@filename = grep /$ss2/, @destfile;
			$document = $word->Documents->Open("$destfulldir/$filename[0]")
   								 || die("Unable to open document ", Win32::OLE->LastError());
			$document->Save;
			$document->Close;		

}



############################################################################################################################
#step9   output summerize information

print "#################################################################\n",
			"                          汇总信息                               \n",
			"#################################################################\n",
			"系统名称：$sys_name \n",
			"需求编号：$req_no\n",
			"需求名称：$req_sum\n",
			"提出日期：$req_date\n",
			"需求提出人：$req_man\n";
			
print "====================================================================\n",
			"在${destfulldir}\t下生成需要提交的文件如下：\n";
			
foreach my $sss (@destfile){
			print $sss,"\n";
		
}

