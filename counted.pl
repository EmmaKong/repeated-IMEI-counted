
#kongqiao 20150808
use Cwd;  
use Time::HiRes qw(gettimeofday);
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
$Win32::OLE::Warn = 3; 
my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');  

my $dir = getcwd;

# ��ȡ���������ļ�
my @allfilePath = <*.xlsx>;
my $filecount = 0;
foreach $path (@allfilePath){
	  	
	 $filePath = $dir."\/".$path;
	 print "����·����$filePath \n";
	 # ��ȡ����
	 $workbook = $Excel->Workbooks->Open($filePath);
	 $workSheet = $workbook->Sheets("ȫ������");
	 #����EXCEL���ݵ�����
	 my $Rowcount = $workSheet->usedrange->rows->count;       #�����Ч���� 
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

# �½� excel�����ڴ洢 ͳ������
my $newworkbook = $Excel->Workbooks->Add(); #�½�һ�������� 
my $newpath = $dir =~ s#/#\\#r;   # ��·���е� ��б�� �滻��б��	 
$file = $newpath.'\\'.'count.xlsx';
$newworkbook->SaveAs($file) or die "Save failer.";

my $newSheet = $newworkbook->Sheets(1);
$newSheet->{name} = "ͳ��";
$newSheet->Range("A1:$Row")->{'value'} = $imei_data;	

$newworkbook->Save();
$newworkbook->Close();

