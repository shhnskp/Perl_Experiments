

use strict;
use warnings;
use Term::ANSIColor;
package Functions;

sub Check_Sheet
{
  my ($InputFiles, $FileName) = @_;
  # Open Excel
  my $ExcelApp = Win32::OLE->GetActiveObject('Excel.Application')
              || do { Win32::OLE->new('Excel.Application', 'Quit')};

  # The file visibility can be set here during script execution
  # 0: Hidden, 1: Shown
  $ExcelApp->{'Visible'} = 0;

  # Navigate upto the Required Worksheet in a Workbook
  my $OpenFile = $InputFiles.$FileName;
  my $Workbook = $ExcelApp->Workbooks->Open($OpenFile);
  my $Worksheet = $Workbook->Worksheets('Header');
  $Worksheet->Activate();

  # Read the Values from the expected cell Range for Status of the sheet and
  # check whether it is open or closed
  my $Data = $Worksheet->Range("E1:E20")->{Value};
  my $Flag = 0;
  for (@$Data)
  {
    for (@$_)
    {
      if (defined $_)
      {
        $Flag = 1 if (($_ eq "Closed") || ($_ eq "closed"));
      }
    }
    last if ($Flag == 1);
  }

  if ($Flag == 1)
  {
    print color("green"), "\n\nThe sheet$OpenFile is CLOSED!!!\n\n", color("reset");
  }
  else
  {
    print color("red"), "\n\nThe sheet $OpenFile is OPEN!!!\n\n", color("reset");
  }

  # $Workbook->SaveAs($OpenFile);
  # Close the Excel Application
  $ExcelApp->close;
}

1;