
#kongqiao 20150808
use Cwd;  
use Time::HiRes qw(gettimeofday);
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
$Win32::OLE::Warn = 3; 
my $Excel = Win32::OLE->GetActiveObject('Excel.Application')|| Win32::OLE->new('Excel.Application', 'Quit');  

my $dir = getcwd;

# ��ȡ���������ļ�
my @allfilePath = <*.xls>;

my $path = $allfilePath[0];  	
my $filePath = $dir."\/".$path;
print "·���� $filePath \n";
# ��ȡ����
my $workbook = $Excel->Workbooks->Open($filePath);
#my $workSheet = $workbook->Sheets("ȫ������");
my $workSheet = $workbook->Sheets(1);
#����EXCEL���ݵ�����
my $Rowcount = $workSheet->usedrange->rows->count;       #�����Ч���� 
$totalRow = $Rowcount + 1;
my $numDRow = Z.$totalRow;
my $allDataArray1 = $workSheet->Range("A1:$numDRow")->{'Value'}; 	 	 
print "aonf $totalRow \n";
my @imei_data1 = 0 x $Rowcount-1;
for($j = 1; $j < $Rowcount; $j++){
	 $imei_data1[$j -1] = $$allDataArray1[$j][3];
} 
$workbook->Save();
$workbook->Close();     


my $path = $allfilePath[1];  	
my $filePath = $dir."\/".$path;
print "·���� $filePath \n";
# ��ȡ����
my $workbook = $Excel->Workbooks->Open($filePath);
#my $workSheet = $workbook->Sheets("ȫ������");
my $workSheet = $workbook->Sheets(1);
#����EXCEL���ݵ�����
my $Rowcount = $workSheet->usedrange->rows->count;       #�����Ч���� 
$totalRow = $Rowcount + 1;
my $numDRow = Z.$totalRow;
my $allDataArray2 = $workSheet->Range("A1:$numDRow")->{'Value'}; 	 
print "aonf $totalRow \n";
my @imei_data2 = 0 x $Rowcount-1;
for($j = 1; $j < $Rowcount; $j++){
	 $imei_data2[$j -1] = $$allDataArray2[$j][3];
} 
$workbook->Save();
$workbook->Close();  

=pod
my $path = $allfilePath[2];  	
my $filePath = $dir."\/".$path;
print "·���� $filePath \n";
# ��ȡ����
my $workbook = $Excel->Workbooks->Open($filePath);
my $workSheet = $workbook->Sheets("ȫ������");
#����EXCEL���ݵ�����
my $Rowcount = $workSheet->usedrange->rows->count;       #�����Ч���� 
$totalRow = $Rowcount + 1;
my $numDRow = X.$totalRow;
my $allDataArray = $workSheet->Range("A1:$numDRow")->{'Value'};  	  
print "aonf $totalRow \n";
my @imei_data3 = 0 x $Rowcount-1;
for($j = 1; $j < $Rowcount; $j++){
	 $imei_data3[$j -1] = $$allDataArray[$j][3];
} 
$workbook->Save();
$workbook->Close();  


my $path = $allfilePath[3];  	
my $filePath = $dir."\/".$path;
print "·���� $filePath \n";
# ��ȡ����
my $workbook = $Excel->Workbooks->Open($filePath);
my $workSheet = $workbook->Sheets("ȫ������");
#����EXCEL���ݵ�����
my $Rowcount = $workSheet->usedrange->rows->count;       #�����Ч���� 
$totalRow = $Rowcount + 1;
my $numDRow = X.$totalRow;
my $allDataArray = $workSheet->Range("A1:$numDRow")->{'Value'};
print "aonf $totalRow \n";
my @imei_data4 = 0 x $Rowcount-1;
for($j = 1; $j < $Rowcount; $j++){
	 $imei_data4[$j -1] = $$allDataArray[$j][3];
}  
$workbook->Save();
$workbook->Close();  

=cut
my @counted_imei = ();
my @counted = 0 x ($#imei_data1+1);

for($i = 0; $i <= $#imei_data1; $i++){
	$counted[$i] ++;
	for($j = 0; $j <= $#imei_data2; $j++){
	
		if($imei_data1[$i] eq $imei_data2[$j]){
			$counted[$i]++;
			#print "ahi $imei_data1[$i] \n";
			last;  # һ�������ظ��������˳�ѭ��������һ��IMEI��һ��excel�е��ظ�ͳ��
			
		}	
	}
=pod	
	for($j = 0; $j <= $#imei_data3; $j++){
	
		if($imei_data1[0] eq $imei_data3[$j]){
			$counted[$i]++;
		#	print "ffg $imei_data1[$i] \n";
			last;
		}	
	}
	
	for($j = 0; $j <= $#imei_data4; $j++){
	
		if($imei_data1[0] eq $imei_data4[$j]){
			$counted[$i]++;
		#	print "fgg $imei_data1[$i] \n";
			last;  # һ�������ظ��������˳�ѭ��������һ��IMEI��һ��excel�е��ظ�ͳ��
		}	
	}
=cut
	
}


# �½� excel�����ڴ洢 ͳ������
my $newworkbook = $Excel->Workbooks->Add(); #�½�һ�������� 
my $newpath = $dir =~ s#/#\\#r;   # ��·���е� ��б�� �滻��б��	 
$file = $newpath.'\\'.'count.xlsx';
$newworkbook->SaveAs($file) or die "Save failer.";
my $newSheet = $newworkbook->Sheets(1);
$newSheet->{name} = "ͳ��";

$rangeEnd = Z.2000; 
#$counted_array = $newSheet->Range("A1:rangeEnd")->{'Value'};

@$counted_array[0] = $$allDataArray1[0];
for($i = 1; $i <= $#imei_data1+1; $i++){
	
	if($counted[$i-1] eq 2){
		push @counted_imei, $imei_data1[$i-1];
		print "fgg $imei_data1[$i] \n";
	  push @$counted_array, $$allDataArray1[$i-1];
	 		
	}
}

my $count = @$counted_array - 1;
print "aoifkb $count \n";

$Num = Z.$count;

$newSheet->Range("A1:$Num")->{'value'} = $counted_array;	

$newworkbook->Save();
$newworkbook->Close();

