unit FileCtrl;

interface
function ReadFile(Filename: string): string;
procedure WriteFile(Filename, Data: string);

implementation

uses
  System.SysUtils,
  System.Classes;

function ReadFile(Filename: string): string;
var
   fileData : TStringList;
   saveLine : String;
   lines, i : Integer;
 begin
   fileData := TStringList.Create;        // Create the TSTringList object
   fileData.LoadFromFile(Filename);     // Load from Testing.txt file

   // Reverse the sequence of lines in the file
   lines := fileData.Count;

   for i := lines-1 downto (lines div 2) do
   begin
     saveLine := fileData[lines-i-1];
     fileData[lines-i-1] := fileData[i];
     fileData[i] := saveLine;
   end;

   // Now display the file
   for i := 0 to lines-1 do
     result := result + fileData[i];

   fileData.SaveToFile(Filename);       // Save the reverse sequence file
 end;

procedure WriteFile(Filename, Data: string);
var
  myFile : TextFile;
begin
  try
    AssignFile(myFile, Filename);
    ReWrite(myFile);
    WriteLn(myFile, Data);
  finally
    CloseFile(myFile);
  end;
end;

end.
