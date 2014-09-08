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
  protected
    procedure AddShortCut(const Key, Value: string);
    procedure AddDelphiShortCut(const Key, RelExePath, RelDocPath: string); overload;
    procedure AddOldDelphiShortCut(const Key, RelExePath, RelDocPath: string); overload;
    function Compress(const Value: string): string;
    function GetExistingPath(const Prefix, RelPath: string): string; overload;
    function GetExistingPath(const Prefix: array of string; const RelPath: string): string; overload;
    procedure InitShortCuts;
    function ReadRegistry(const AName: string): string;
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
  if Value > '' then begin
    Shortcuts.Values[Key] := Value;
  end;
end;

procedure TPathCompressor.AddDelphiShortCut(const Key, RelExePath, RelDocPath: string);
begin
  AddShortCut(Key, GetExistingPath(PathDelphi, RelExePath));
  AddShortCut(Key + 'BPL', GetExistingPath([PathMyDocs, PathSharedDocs], RelDocPath));
end;

procedure TPathCompressor.AddOldDelphiShortCut(const Key, RelExePath, RelDocPath: string);
begin
  AddShortCut(Key, GetExistingPath(PathDelphi, RelExePath));
  AddShortCut(Key + 'BPL', GetExistingPath(PathDelphi, RelDocPath));
end;

function TPathCompressor.Compress(const Value: string): string;
var
  S: string;
  N: Integer;
  L: Integer;
  I: Integer;
begin
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
    result := Value.Remove(0, L).Insert(0, '%' + Shortcuts.Names[N] + '%');
    Variables.Append(Shortcuts[N]);
  end
  else begin
    result := Value;
  end;
end;

procedure TPathCompressor.Execute(out OldLength, NewLength: Integer);
var
  I: Integer;
begin
  OldLength := OrgLength;
  for I := 0 to Path.Count - 1 do begin
    Path[I] := Compress(Path[I]);
  end;
  StoreVariables;
  NewLength := StorePath;
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
  { Allgemeine Programmpfade }
  AddShortCut('PF', PathProgramFiles);
  AddShortCut('PF86', PathProgramFilesX86);

  { SQL Server }
  AddShortCut('SQL', GetExistingPath(PathProgramFiles, '\Microsoft SQL Server'));
  AddShortCut('SQL86', GetExistingPath(PathProgramFilesX86, '\Microsoft SQL Server'));

  { TODO : Pfade für fehlende Delphi-Versionen ergänzen}

  { Delphi 7 }
  AddOldDelphiShortCut('D7', '\Borland\Delphi7', '\Borland\Delphi7\Projects\Bpl');

  { Delphi ab D2007 }
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
    Path.DelimitedText := S;
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
  Writeln(Name, '=', Value);
  Exit;
  {$ENDIF}
  if Value.Contains('%') then
    Registry.WriteExpandString(Name, Value)
  else
    Registry.WriteString(Name, Value);
end;

procedure Main;
var
  instance: TPathCompressor;
  lNew: Integer;
  lOld: Integer;
  shortCutFileName: string;
begin
  instance := TPathCompressor.Create;
  try
    if FindCmdLineSwitch('f', shortCutFileName) then begin
      instance.LoadShortCuts(shortCutFileName);
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
  Write('Press <Enter> to continue...');
  Readln;
  {$ENDIF}
end.
