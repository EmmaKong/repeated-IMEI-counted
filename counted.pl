
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
foreach $path (@allfilePath){
	  	
	 $filePath = $dir."\/".$path;
	 print "����·����$filePath \n";
	  	     
}
my $filecount = @allfilePath;
print "$filecount \n";