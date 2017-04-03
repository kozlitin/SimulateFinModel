#!perl 

use strict;

use Win32::OLE;
use Getopt::Long;
use File::Basename;
use File::Spec;
use Cwd;

my $ExcelFileName;

GetOptions (  "excelfilename=s"  => \$ExcelFileName
			  ) or &{ 
					print_usage();
					exit(1)
					};
					
if(dirname($ExcelFileName) =~ /^\.$/) {
	$ExcelFileName = File::Spec->catfile( getcwd , $ExcelFileName );
}					
					
my $ExcelOle = Win32::OLE->new('Excel.Application', 'Quit');
my $ExcelBookOle = $ExcelOle->Workbooks->Open($ExcelFileName);

if (!$ExcelBookOle) {
	print "Can not open workbook $ExcelFileName\n";
	$ExcelOle->Quit();
	$ExcelOle = undef;
	exit(1);
}					

my ($SheetWithVariables, $SheetWithData, $SheetWithDataStat, $SheetWithSummary, $SheetDetailed);
					
for (my $i=1; $i <= $ExcelBookOle->Sheets->{Count}; $i++ ) {

	my $sheet = $ExcelBookOle->Worksheets($i);

	$SheetWithVariables = $sheet if $sheet->{Name} =~ /Variables/;
	
	$SheetWithData = $sheet if $sheet->{Name} =~ /^Data$/;
	
	$SheetWithDataStat = $sheet if $sheet->{Name} =~ /^Data \(Stat\)$/;
	
	$SheetWithSummary = $sheet if $sheet->{Name} =~ /^Summary$/;
	
	$SheetDetailed = $sheet if $sheet->{Name} =~ /^Detailed$/;

}	
	
my %Variables;	
	
foreach my $row (2 .. $SheetWithVariables->Cells->SpecialCells(11)->{Row}) {

	my $Version = $SheetWithVariables->Cells($row,5)->{Value};
	
	next unless $Version =~ /\d+/;

	$Variables{$Version} = {
		DC1RackPriceDec => $SheetWithVariables->Cells($row,6)->{Value},
		DC1NewSalesPrice => $SheetWithVariables->Cells($row,7)->{Value},
		DC5RackPrice => $SheetWithVariables->Cells($row,8)->{Value},
		DC5RackSales => $SheetWithVariables->Cells($row,9)->{Value},
		CloudGrowth => $SheetWithVariables->Cells($row,10)->{Value},
		CloudPriceDec => $SheetWithVariables->Cells($row,11)->{Value},
		SalaryInc => $SheetWithVariables->Cells($row,12)->{Value},
		MarketingInc => $SheetWithVariables->Cells($row,13)->{Value},
	};

}
	
foreach my $key (sort {$a <=> $b} (keys %Variables)) {
	print $key, "\t", join ("\t",@{[%{$Variables{$key}}]}), "\n";
}	
	
$SheetWithDataStat->Range("A2:J50000")->ClearContents();
	
my $DataStatBlockOffset = 0;	
	
foreach my $key (sort {$a <=> $b} (keys %Variables)) {
	
	print "Version: $key\n";
	
	$SheetWithSummary->Cells(2,3)->{Value} = $Variables{$key}->{DC1RackPriceDec};
	$SheetWithSummary->Cells(3,3)->{Value} = $Variables{$key}->{DC1NewSalesPrice};
	$SheetWithSummary->Cells(4,3)->{Value} = $Variables{$key}->{DC5RackPrice};
	$SheetWithSummary->Cells(5,3)->{Value} = $Variables{$key}->{DC5RackSales};
	$SheetWithSummary->Cells(6,3)->{Value} = $Variables{$key}->{CloudGrowth};
	$SheetWithSummary->Cells(7,3)->{Value} = $Variables{$key}->{CloudPriceDec};
	$SheetWithSummary->Cells(8,3)->{Value} = $Variables{$key}->{SalaryInc};
	$SheetWithSummary->Cells(9,3)->{Value} = $Variables{$key}->{MarketingInc};
	
	$SheetDetailed->Calculate();
	$SheetWithData->Calculate();
	
	for (my $sr=2; $sr<=39; $sr++) {
		$SheetWithDataStat->Cells($sr+$DataStatBlockOffset*38, 1)->{Value} = $key;
		for (my $sc=1; $sc<=9; $sc++) {
			$SheetWithDataStat->Cells($sr+$DataStatBlockOffset*38, $sc+1)->{Value} = $SheetWithData->Cells($sr,$sc)->{Value};	
		}	
	}
	
	$DataStatBlockOffset++;
}		
	
$ExcelBookOle->save();

print "Workbook was saved.\n";	
	
$ExcelOle->Quit();
$ExcelOle = undef;	
	
######################################################
# Subroutines
######################################################					
					
sub print_usage {
	print "Usage: perl $0 --excelfilename=excel_file_name\n";
}					