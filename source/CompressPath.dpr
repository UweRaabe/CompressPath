program CompressPath;

{$APPTYPE CONSOLE}
{$R *.res}

uses
  System.SysUtils,
  System.Classes,
  System.Win.Registry,
  Winapi.Windows,
  Winapi.Messages;

type
  TPathCompressor = class
  private
    FOrgLength: Integer;
    FPath: TStringList;
    FRegistry: TRegistry;
    FShortcuts: TStringList;
    FVariables: TStringList;
  protected
    procedure InitShortCuts;
    function ReadRegistry(const AName: string): string;
    procedure WriteRegistry(const Name, Value: string);
  public
    constructor Create;
    destructor Destroy; override;
    function Compress(const Value: string): string;
    procedure Execute(out OldLength, NewLength: Integer);
    function LoadPath(UserEnvironment: Boolean): Boolean;
    procedure NotifyChanges;
    function StorePath: Integer;
    procedure StoreVariables;
    property OrgLength: Integer read FOrgLength;
    property Path: TStringList read FPath;
    property Registry: TRegistry read FRegistry;
    property Shortcuts: TStringList read FShortcuts;
    property Variables: TStringList read FVariables;
  end;

constructor TPathCompressor.Create;
begin
  inherited Create;
  FShortcuts := TStringList.Create;
  FVariables := TStringList.Create();
  FPath := TStringList.Create;
  FPath.Delimiter := ';';
  FPath.StrictDelimiter := true;
  {$IFDEF DEBUG}
  FRegistry := TRegistry.Create(KEY_READ);
  {$ELSE}
  FRegistry := TRegistry.Create();
  {$ENDIF}
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
    Variables.Add(Shortcuts[N]);
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

procedure TPathCompressor.InitShortCuts;
begin
  { TODO : Adapt to current operating system }
  { Common Pathes under Windows 7/8 (x64) }
  Shortcuts.Values['PF'] := 'C:\Program Files';
  Shortcuts.Values['PF86'] := 'C:\Program Files (x86)';
  Shortcuts.Values['DOCS'] := 'C:\Users\Public\Documents';
  Shortcuts.Values['SQL'] := 'C:\Program Files\Microsoft SQL Server';
  Shortcuts.Values['SQL86'] := 'C:\Program Files (x86)\Microsoft SQL Server';
  Shortcuts.Values['D7'] := 'C:\Program Files (x86)\Borland\Delphi7';
  Shortcuts.Values['BPL7'] := 'C:\Program Files (x86)\Borland\Delphi7\Projects\Bpl';
  Shortcuts.Values['D2007'] := 'C:\Program Files (x86)\CodeGear\RAD Studio\5.0';
  Shortcuts.Values['BPL2007'] := 'C:\Users\Public\Documents\RAD Studio\5.0\Bpl';
  Shortcuts.Values['D2009'] := 'C:\Program Files (x86)\CodeGear\RAD Studio\6.0';
  Shortcuts.Values['BPL2009'] := 'C:\Users\Public\Documents\RAD Studio\6.0\Bpl';
  Shortcuts.Values['D2010'] := 'C:\Program Files (x86)\Embarcadero\RAD Studio\7.0';
  Shortcuts.Values['BPL2010'] := 'C:\Users\Public\Documents\RAD Studio\7.0\Bpl';
  Shortcuts.Values['XE'] := 'C:\Program Files (x86)\Embarcadero\RAD Studio\8.0';
  Shortcuts.Values['BPLXE'] := 'C:\Users\Public\Documents\RAD Studio\8.0\Bpl';
  Shortcuts.Values['XE2'] := 'C:\Program Files (x86)\Embarcadero\RAD Studio\9.0';
  Shortcuts.Values['BPLXE2'] := 'C:\Users\Public\Documents\RAD Studio\9.0\Bpl';
  Shortcuts.Values['XE3'] := 'C:\Program Files (x86)\Embarcadero\RAD Studio\10.0';
  Shortcuts.Values['BPLXE3'] := 'C:\Users\Public\Documents\RAD Studio\10.0\Bpl';
  Shortcuts.Values['XE4'] := 'C:\Program Files (x86)\Embarcadero\RAD Studio\11.0';
  Shortcuts.Values['BPLXE4'] := 'C:\Users\Public\Documents\RAD Studio\11.0\Bpl';
  Shortcuts.Values['XE5'] := 'C:\Program Files (x86)\Embarcadero\RAD Studio\12.0';
  Shortcuts.Values['BPLXE5'] := 'C:\Users\Public\Documents\RAD Studio\12.0\Bpl';
  Shortcuts.Values['XE6'] := 'C:\Program Files (x86)\Embarcadero\Studio\14.0';
  Shortcuts.Values['BPLXE6'] := 'C:\Users\Public\Documents\Embarcadero\Studio\14.0\Bpl';
  Shortcuts.Values['XE7'] := 'C:\Program Files (x86)\Embarcadero\Studio\15.0';
  Shortcuts.Values['BPLXE7'] := 'C:\Users\Public\Documents\Embarcadero\Studio\15.0\Bpl';
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

procedure TPathCompressor.NotifyChanges;
{ Sending a WM_SETTINGCHANGE message to all top level windows. Otherwise the new environment variables
  will only be visible after logoff/logon. }
var
  env: PChar;
  lparam: NativeInt;
begin
  env := 'Environment';
  lparam := NativeInt(env);
  SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0, lParam, SMTO_ABORTIFHUNG, 5000, nil);
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
begin
  instance := TPathCompressor.Create;
  try
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
