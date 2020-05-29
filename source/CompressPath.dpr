program CompressPath;

{$APPTYPE CONSOLE}
{$R *.res}

uses
  Winapi.Windows,
  Winapi.Messages,
  System.SysUtils,
  System.Classes,
  System.IOUtils,
  System.Win.Registry;

type
  TPathCompressor = class
  private
    FOrgLength: Integer;
    FPath: TStringList;
    FPathDelphi: string;
    FPathMyDocs: string;
    FPathProgramFiles: string;
    FPathProgramFilesX86: string;
    FPathSharedDocs: string;
    FRegistry: TRegistry;
    FShortcuts: TStringList;
    FVariables: TStringList;
    function ExpandEnvironmentVars(const Value: string): string;
  protected
    procedure AddShortCut(const Key, Value: string);
    procedure AddDelphiShortCut(const Key, RelExePath, RelDocPath: string); overload;
    procedure AddOldDelphiShortCut(const Key, RelExePath, RelDocPath: string); overload;
    function Compress(const Value: string): string;
    function GetExistingPath(const Prefix, RelPath: string): string; overload;
    function GetExistingPath(const Prefix: array of string; const RelPath: string): string; overload;
    procedure InitShortCuts;
    function ReadRegistry(const AName: string): string;
    procedure RemoveDuplicates;
    function StorePath: Integer;
    procedure StoreVariables;
    procedure WriteRegistry(const Name, Value: string);
    property OrgLength: Integer read FOrgLength;
    property Path: TStringList read FPath;
    property PathDelphi: string read FPathDelphi;
    property PathMyDocs: string read FPathMyDocs;
    property PathProgramFiles: string read FPathProgramFiles;
    property PathProgramFilesX86: string read FPathProgramFilesX86;
    property PathSharedDocs: string read FPathSharedDocs;
    property Registry: TRegistry read FRegistry;
    property Shortcuts: TStringList read FShortcuts;
    property Variables: TStringList read FVariables;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Execute(out OldLength, NewLength: Integer);
    function LoadPath(UserEnvironment: Boolean): Boolean;
    procedure LoadShortCuts(const AFileName: string);
    procedure StoreShortCuts(const AFileName: string);
    procedure NotifyChanges;
  end;

constructor TPathCompressor.Create;
begin
  inherited Create;
  FShortcuts := TStringList.Create;
  FVariables := TStringList.Create;
  FVariables.Sorted := true;
  FVariables.Duplicates := dupIgnore;
  FPath := TStringList.Create;
  FPath.Delimiter := ';';
  FPath.StrictDelimiter := true;
  {$IFDEF DEBUG}
  { It is quite unlikely that we have write permission to HKEY_LOCAL_MACHINE when we are running from inside the IDE.
    In addition it doesn't make much sense to run in DEBUG mode outside of the IDE. Hence this restriction. There are
    other places where we check for DEBUG mode. In case you get unexpected results, please have a look for those. }
  FRegistry := TRegistry.Create(KEY_READ);
  {$ELSE}
  FRegistry := TRegistry.Create();
  {$ENDIF}
  FPathMyDocs := TPath.GetDocumentsPath;
  FPathSharedDocs := TPath.GetSharedDocumentsPath;
  FPathProgramFilesX86 := GetEnvironmentVariable('ProgramFiles(x86)');
  if FPathProgramFilesX86 = '' then begin
    FPathProgramFiles := GetEnvironmentVariable('ProgramFiles');
    FPathDelphi := FPathProgramFiles;
  end
  else begin
    FPathProgramFiles := GetEnvironmentVariable('ProgramW6432');
    FPathDelphi := FPathProgramFilesX86;
  end;
  InitShortCuts;
end;

destructor TPathCompressor.Destroy;
begin
  FRegistry.Free;
  FPath.Free;
  FVariables.Free;
  FShortcuts.Free;
  inherited Destroy;
end;

procedure TPathCompressor.AddShortCut(const Key, Value: string);
begin
  { Shortcuts are only usefull when they contain some data. }
  if Value > '' then begin
    Shortcuts.Values[Key] := Value;
  end;
end;

procedure TPathCompressor.AddDelphiShortCut(const Key, RelExePath, RelDocPath: string);
begin
  AddShortCut(Key, GetExistingPath(PathDelphi, RelExePath));
  { Depending on the type of installation (current user or all users), the BPLs are located
    in different documents folders. Here we first check inside the personal folder and
    if that doesn't exist, we take the public documents folder.
  }
  AddShortCut(Key + 'BPL', GetExistingPath([PathMyDocs, PathSharedDocs], RelDocPath));
end;

procedure TPathCompressor.AddOldDelphiShortCut(const Key, RelExePath, RelDocPath: string);
begin
  AddShortCut(Key, GetExistingPath(PathDelphi, RelExePath));
  { In older Delphi versions the BPL folder resides below the Delphi folder. }
  AddShortCut(Key + 'BPL', GetExistingPath(PathDelphi, RelDocPath));
end;

function TPathCompressor.Compress(const Value: string): string;
var
  S: string;
  N: Integer;
  L: Integer;
  I: Integer;
begin
  { Find the Shortcut that will replace the longest string from the beginning of Value.
    Ideally it will replace the whole string.
  }
  N := -1;
  L := 0;
  for I := 0 to Shortcuts.Count - 1 do begin
    S := Shortcuts.ValueFromIndex[I];
    if Value.StartsWith(S, true) and (L < S.Length) then begin
      N := I;
      L := S.Length;
    end;
  end;

  if L > 0 then begin
    { Replace the Shortcut and remember to add a similar variable. }
    result := Value.Remove(0, L).Insert(0, '%' + Shortcuts.Names[N] + '%');
    Variables.Append(Shortcuts[N]);
  end
  else begin
    { Nothing to compress. }
    result := Value;
  end;
end;

procedure TPathCompressor.Execute(out OldLength, NewLength: Integer);
{ Returns the old and new length of the PATH variable. }
var
  I: Integer;
begin
  Variables.Clear;
  OldLength := OrgLength;
  for I := 0 to Path.Count - 1 do begin
    Path[I] := Compress(Path[I]);
  end;
  RemoveDuplicates;
  StoreVariables;
  NewLength := StorePath;
end;

function TPathCompressor.ExpandEnvironmentVars(const Value: string): string;
var
  buffer: TCharArray;
  res: Cardinal;
begin
  Result := Value;
  SetLength(Buffer, 32*1024);
  res := ExpandEnvironmentStrings(PChar(Value), @Buffer[0], Length(Buffer) - 1);
  if res > 0 then begin
    Result := PChar(@Buffer[0]);
  end;
end;

function TPathCompressor.GetExistingPath(const Prefix: array of string; const
    RelPath: string): string;
var
  S: string;
begin
  for S in Prefix do begin
    Result := S + RelPath;
    if TDirectory.Exists(Result) then Exit;
  end;
  Result := '';
end;

function TPathCompressor.GetExistingPath(const Prefix, RelPath: string): string;
begin
  Result := GetExistingPath([Prefix], RelPath);
end;

procedure TPathCompressor.InitShortCuts;
begin
  { Program Files }
  AddShortCut('PF', PathProgramFiles);
  AddShortCut('PF86', PathProgramFilesX86);

  { SQL Server }
  AddShortCut('SQL', GetExistingPath(PathProgramFiles, '\Microsoft SQL Server'));
  AddShortCut('SQL86', GetExistingPath(PathProgramFilesX86, '\Microsoft SQL Server'));

  { TODO : Add folders for missing Delphi versions}

  { Delphi 7 }
  AddOldDelphiShortCut('D7', '\Borland\Delphi7', '\Borland\Delphi7\Projects\Bpl');

  { Delphi D2007+ }
  AddDelphiShortCut('D2007', '\CodeGear\RAD Studio\5.0', '\RAD Studio\5.0\Bpl');
  AddDelphiShortCut('D2009', '\CodeGear\RAD Studio\6.0', '\RAD Studio\6.0\Bpl');
  AddDelphiShortCut('D2010', '\Embarcadero\RAD Studio\7.0', '\RAD Studio\7.0\Bpl');
  AddDelphiShortCut('XE', '\Embarcadero\RAD Studio\8.0', '\RAD Studio\8.0\Bpl');
  AddDelphiShortCut('XE2', '\Embarcadero\RAD Studio\9.0', '\RAD Studio\9.0\Bpl');
  AddDelphiShortCut('XE3', '\Embarcadero\RAD Studio\10.0', '\RAD Studio\10.0\Bpl');
  AddDelphiShortCut('XE4', '\Embarcadero\RAD Studio\11.0', '\RAD Studio\11.0\Bpl');
  AddDelphiShortCut('XE5', '\Embarcadero\RAD Studio\12.0', '\RAD Studio\12.0\Bpl');
  AddDelphiShortCut('XE6', '\Embarcadero\Studio\14.0', '\Embarcadero\Studio\14.0\Bpl');
  AddDelphiShortCut('XE7', '\Embarcadero\Studio\15.0', '\Embarcadero\Studio\15.0\Bpl');
  AddDelphiShortCut('XE8', '\Embarcadero\Studio\16.0', '\Embarcadero\Studio\16.0\Bpl');
  AddDelphiShortCut('DX', '\Embarcadero\Studio\17.0', '\Embarcadero\Studio\17.0\Bpl');
  AddDelphiShortCut('DX1', '\Embarcadero\Studio\18.0', '\Embarcadero\Studio\18.0\Bpl');
  AddDelphiShortCut('DX2', '\Embarcadero\Studio\19.0', '\Embarcadero\Studio\19.0\Bpl');
  AddDelphiShortCut('DX3', '\Embarcadero\Studio\20.0', '\Embarcadero\Studio\20.0\Bpl');
  AddDelphiShortCut('DX4', '\Embarcadero\Studio\21.0', '\Embarcadero\Studio\21.0\Bpl');
end;

function TPathCompressor.LoadPath(UserEnvironment: Boolean): Boolean;
const
  cKeyHKLM = 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment';
  cKeyHKCU = 'Environment';
var
  S: string;
  RegKey: string;
begin
  Path.Clear;
  if UserEnvironment then begin
    Registry.RootKey := HKEY_CURRENT_USER;
    RegKey := cKeyHKCU;
  end
  else begin
    Registry.RootKey := HKEY_LOCAL_MACHINE;
    RegKey := cKeyHKLM;
  end;
  result := Registry.OpenKey(RegKey, false);
  if result then begin
    S := ReadRegistry('Path');
    Path.DelimitedText := ExpandEnvironmentVars(S);
    FOrgLength := S.Length;
  end;
end;

procedure TPathCompressor.LoadShortCuts(const AFileName: string);
begin
  Shortcuts.LoadFromFile(AFileName);
end;

procedure TPathCompressor.StoreShortCuts(const AFileName: string);
begin
  Shortcuts.SaveToFile(AFileName);
end;

procedure TPathCompressor.NotifyChanges;
{ Sending a WM_SETTINGCHANGE message to all top level windows. Otherwise the new environment variables
  will only be visible after logoff/logon. }
begin
  {$IFDEF DEBUG}
  Exit;
  {$ENDIF}
  SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, NativeInt(PChar('Environment')), SMTO_ABORTIFHUNG, 5000, nil);
end;

function TPathCompressor.ReadRegistry(const AName: string): string;
begin
  result := Registry.ReadString(AName);
end;

procedure TPathCompressor.RemoveDuplicates;
var
  lst: TStringList;
  S: string;
  I: Integer;
  N: Integer;
begin
  lst := TStringList.Create;
  try
    lst.Sorted := true;
    lst.Duplicates := dupIgnore;
    for S in Path do begin
      lst.Add(S);
    end;
    I := 0;
    while I < Path.Count do begin
      if lst.Find(Path[I], N) then begin
        lst.Delete(N);
        Inc(I);
      end
      else begin
        Path.Delete(I);
      end;
    end;
  finally
    lst.Free;
  end;
end;

function TPathCompressor.StorePath: Integer;
var
  S: string;
begin
  S := Path.DelimitedText;
  WriteRegistry('Path', S);
  result := S.Length;
end;

procedure TPathCompressor.StoreVariables;
var
  I: Integer;
begin
  Variables.Sort;
  for I := 0 to Variables.Count - 1 do begin
    WriteRegistry(Variables.Names[I], Variables.ValueFromIndex[I]);
  end;
end;

procedure TPathCompressor.WriteRegistry(const Name, Value: string);
begin
  {$IFDEF DEBUG}
  { In DEBUG mode only show the changes that would be made to the registry.
  }
  Writeln(Name, '=', Value);
  Exit;
  {$ENDIF}
  if Value.Contains('%') then
    Registry.WriteExpandString(Name, Value)
  else
    Registry.WriteString(Name, Value);
end;

procedure ShowHelp;
var
  exeName: string;
begin
  exeName := TPath.GetFileNameWithoutExtension(ParamStr(0));
  Writeln('Compresses the system and user PATH variable using environment variables.');
  Writeln;
  Writeln(exeName, ' [/F importfile] [/X exportfile]');
  Writeln;
  Writeln('  /F importfile'.PadRight(20), 'load shortcuts from <importfile>');
  Writeln('  /X exportfile'.PadRight(20), 'store shortcuts into <exportfile>');
  Writeln;
  Writeln('You must have Administrator rights to change the system PATH variable!');
  Writeln;
end;

procedure Main;
var
  instance: TPathCompressor;
  lNew: Integer;
  lOld: Integer;
  shortCutFileName: string;
begin
  if FindCmdLineSwitch('?') then begin
    ShowHelp;
    Exit;
  end;
  instance := TPathCompressor.Create;
  try
    if FindCmdLineSwitch('f', shortCutFileName) then begin
      instance.LoadShortCuts(shortCutFileName);
    end;
    if FindCmdLineSwitch('x', shortCutFileName) then begin
      instance.StoreShortCuts(shortCutFileName);
    end;
    Write('HKEY_LOCAL_MACHINE: ');
    if instance.LoadPath(false) then begin
      instance.Execute(lOld, lNew);
      Writeln(lOld, ' ==> ', lNew);
    end
    else begin
      Writeln('Could not load path. Try running as administrator.');
    end;
    Write('HKEY_CURRENT_USER: ');
    if instance.LoadPath(true) then begin
      instance.Execute(lOld, lNew);
      Writeln(lOld, ' ==> ', lNew);
    end
    else begin
      Writeln('Could not load path. Try running as administrator.');
    end;
    Write('notifying...');
    instance.NotifyChanges;
    Writeln;
  finally
    instance.Free;
  end;
end;

begin
  try
    Main;
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;
  {$IFDEF DEBUG}
  { When started from within the IDE we want to see the output. }
  Write('Press <Enter> to continue...');
  Readln;
  {$ENDIF}
end.
