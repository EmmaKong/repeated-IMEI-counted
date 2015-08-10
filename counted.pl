
#kongqiao 20150808
use Cwd;  
use Time::HiRes qw(gettimeofday);
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
$Win32::OLE::Warn = 3; 
my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');  

my $dir = getcwd;

# 获取所有数据文件
my @allfilePath = <*.xlsx>;
my $filecount = 0;
foreach $path (@allfilePath){
	  	
	 $filePath = $dir."\/".$path;
	 print "数据路径：$filePath \n";
	 # 获取数据
	 $workbook = $Excel->Workbooks->Open($filePath);
	 $workSheet = $workbook->Sheets("全部数据");
	 #读出EXCEL数据到数组
	 my $Rowcount = $workSheet->usedrange->rows->count;       #最大有效行数 
	 my $totolRow = $Rowcount + 1;
	 my $numDRow = X.$totolRow;
	 my $allDataArray = $workSheet->Range("A1:$numDRow")->{'Value'};
	 
	 for ($i = 0; $i< $totolRow; $i++){
	 	 $imei_data[$filecount][$i] = $allDataArray[$i][3];
	 }
	 $filecount++;	  	
	 
	 
	 $workbook->Save();
	 $workbook->Close();     
}
$Row = X.($filecount - 1);

# 新建 excel，用于存储 统计数据
my $newworkbook = $Excel->Workbooks->Add(); #新建一个工作簿 
my $newpath = $dir =~ s#/#\\#r;   # 将路径中的 反斜杠 替换成斜杠	 
$file = $newpath.'\\'.'count.xlsx';
$newworkbook->SaveAs($file) or die "Save failer.";

my $newSheet = $newworkbook->Sheets(1);
$newSheet->{name} = "统计";
$newSheet->Range("A1:$Row")->{'value'} = $imei_data;	

$newworkbook->Save();
$newworkbook->Close();

