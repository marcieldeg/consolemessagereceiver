unit UConsoleMessageReceiver;

interface

uses
  Windows, SysUtils, Variants, Classes, ActiveX, StdCtrls, SHDocVw, MSHTML, Controls;

type
{$IF NOT DECLARED(IDeveloperConsoleMessageReceiver)}
  _DEV_CONSOLE_MESSAGE_LEVEL = TOleEnum;

const
  DCML_INFORMATIONAL = $00000000;
  DCML_WARNING = $00000001;
  DCML_ERROR = $00000002;
  DEV_CONSOLE_MESSAGE_LEVEL_Max = $7FFFFFFF;

type
  IDeveloperConsoleMessageReceiver = interface(IUnknown)
    ['{30510808-98B5-11CF-BB82-00AA00BDCE0B}']
    function write(source: PWideChar; level: _DEV_CONSOLE_MESSAGE_LEVEL; messageId: SYSINT;
      messageText: PWideChar): HResult; stdcall;
    function WriteWithUrl(source: PWideChar; level: _DEV_CONSOLE_MESSAGE_LEVEL; messageId: SYSINT;
      messageText: PWideChar; fileUrl: PWideChar): HResult; stdcall;
    function WriteWithUrlAndLine(source: PWideChar; level: _DEV_CONSOLE_MESSAGE_LEVEL; messageId: SYSINT;
      messageText: PWideChar; fileUrl: PWideChar; line: LongWord): HResult; stdcall;
    function WriteWithUrlLineAndColumn(source: PWideChar; level: _DEV_CONSOLE_MESSAGE_LEVEL; messageId: SYSINT;
      messageText: PWideChar; fileUrl: PWideChar; line: LongWord; column: LongWord): HResult; stdcall;
  end;
{$IFEND}

  TConsoleMessageReceiver = class(TCustomMemo, IDeveloperConsoleMessageReceiver)
  private
    FWebBrowser: TWebBrowser;
    FShowInfo: Boolean;
    procedure RegisterMessageReceiver(ARegister: Boolean);
    function Write(source: PWideChar; level: _DEV_CONSOLE_MESSAGE_LEVEL; messageId: SYSINT;
      messageText: PWideChar): HResult; stdcall;
    function WriteWithUrl(source: PWideChar; level: _DEV_CONSOLE_MESSAGE_LEVEL; messageId: SYSINT;
      messageText: PWideChar; fileUrl: PWideChar): HResult; stdcall;
    function WriteWithUrlAndLine(source: PWideChar; level: _DEV_CONSOLE_MESSAGE_LEVEL; messageId: SYSINT;
      messageText: PWideChar; fileUrl: PWideChar; line: LongWord): HResult; stdcall;
    function WriteWithUrlLineAndColumn(source: PWideChar; level: _DEV_CONSOLE_MESSAGE_LEVEL; messageId: SYSINT;
      messageText: PWideChar; fileUrl: PWideChar; line: LongWord; column: LongWord): HResult; stdcall;
    procedure SetWebBrowser(AWebBrowser: TWebBrowser);
  public
    constructor Create(AOwner: TComponent); override;
  published
    property WebBrowser: TWebBrowser read FWebBrowser write SetWebBrowser;
    property ShowInfo: Boolean read FShowInfo write FShowInfo default True;
    property Align;
    property Alignment;
    property Anchors;
    property BevelEdges;
    property BevelInner;
    property BevelKind default bkNone;
    property BevelOuter;
    property BiDiMode;
    property BorderStyle;
    property Color;
    property Constraints;
    property Ctl3D;
    property DoubleBuffered;
    property Enabled;
    property Font;
    property HideSelection;
    property ImeMode;
    property ImeName;
    property OEMConvert;
    property ParentBiDiMode;
    property ParentColor;
    property ParentCtl3D;
    property ParentDoubleBuffered;
    property ParentFont;
    property ParentShowHint;
    property PopupMenu;
    property ScrollBars default ssVertical;
    property ShowHint;
    property TabOrder;
    property TabStop;
    property Touch;
    property Visible;
    property WantReturns;
    property WantTabs;
    property WordWrap;
    property OnChange;
    property OnClick;
    property OnContextPopup;
    property OnDblClick;
    property OnDragDrop;
    property OnDragOver;
    property OnEndDock;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
    property OnGesture;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnMouseActivate;
    property OnMouseDown;
    property OnMouseEnter;
    property OnMouseLeave;
    property OnMouseMove;
    property OnMouseUp;
    property OnStartDock;
    property OnStartDrag;
  end;

procedure Register;

implementation

procedure Register;
begin
  Classes.RegisterComponents('Internet', [TConsoleMessageReceiver]);
end;

{TConsoleMessageReceiver}
constructor TConsoleMessageReceiver.Create(AOwner: TComponent);
begin
  inherited;
  Self.FShowInfo := True;
  Self.ScrollBars := ssVertical;
  Self.ReadOnly := True;
  Self.ParentFont := False;
  Self.Font.Name := 'Courier';
end;

procedure TConsoleMessageReceiver.RegisterMessageReceiver(ARegister: Boolean);
const
  IDM_ADDCONSOLEMESSAGERECEIVER = 3800;
  IDM_REMOVECONSOLEMESSAGERECEIVER = 3801;
  CGID_MSHTML: TGUID = '{DE4BA900-59CA-11CF-9592-444553540000}';
var
  Comm: IOleCommandTarget;
  Action: Cardinal;
  OutVar: OleVariant;
begin
  if csDesigning in ComponentState then
    Exit;

  if ARegister then
    Action := IDM_ADDCONSOLEMESSAGERECEIVER
  else
    Action := IDM_REMOVECONSOLEMESSAGERECEIVER;

  if Assigned(Self.FWebBrowser) then
  begin
    OutVar := EmptyParam;
    if not Assigned(Self.FWebBrowser.Document) then
      Self.FWebBrowser.Navigate('about:blank');
    if Supports(Self.FWebBrowser.Document, IOleCommandTarget, Comm) then
      Comm.Exec(@CGID_MSHTML, Action, OLECMDEXECOPT_DODEFAULT, IDeveloperConsoleMessageReceiver(Self), OutVar);
  end;
end;

procedure TConsoleMessageReceiver.SetWebBrowser(AWebBrowser: TWebBrowser);
begin
  if Self.FWebBrowser <> AWebBrowser then
  begin
    Self.Lines.Clear;
    RegisterMessageReceiver(False);
    Self.FWebBrowser := AWebBrowser;
    RegisterMessageReceiver(True);
  end;
end;

function TConsoleMessageReceiver.Write(source: PWideChar; level: _DEV_CONSOLE_MESSAGE_LEVEL; messageId: SYSINT;
  messageText: PWideChar): HResult;
const
  MSG_MODEL = '%s'#9'Code %d - %s';
var
  LevelType: String;
begin
  Result := S_OK;
  case level of
    DCML_INFORMATIONAL:
      begin
        if not Self.FShowInfo then
          Exit;
        LevelType := 'INFO';
      end;
    DCML_WARNING:
      LevelType := 'WARN';
    DCML_ERROR:
      LevelType := 'ERROR';
  end;
  Self.Lines.Add(Format(MSG_MODEL, [LevelType, messageId, String(messageText)]));
end;

function TConsoleMessageReceiver.WriteWithUrl(source: PWideChar; level: _DEV_CONSOLE_MESSAGE_LEVEL; messageId: SYSINT;
  messageText, fileUrl: PWideChar): HResult;
const
  MSG_MODEL = '%s'#9'Code %d - %s'#13#10#9'> at %s';
var
  LevelType: String;
begin
  Result := S_OK;
  case level of
    DCML_INFORMATIONAL:
      begin
        if not Self.FShowInfo then
          Exit;
        LevelType := 'INFO';
      end;
    DCML_WARNING:
      LevelType := 'WARN';
    DCML_ERROR:
      LevelType := 'ERROR';
  end;
  Self.Lines.Add(Format(MSG_MODEL, [LevelType, messageId, String(messageText), String(fileUrl)]));
end;

function TConsoleMessageReceiver.WriteWithUrlAndLine(source: PWideChar; level: _DEV_CONSOLE_MESSAGE_LEVEL;
  messageId: SYSINT; messageText, fileUrl: PWideChar; line: LongWord): HResult;
const
  MSG_MODEL = '%s'#9'Code %d - %s'#13#10#9'> at %s'#13#10#9'> in line %d';
var
  LevelType: String;
begin
  Result := S_OK;
  case level of
    DCML_INFORMATIONAL:
      begin
        if not Self.FShowInfo then
          Exit;
        LevelType := 'INFO';
      end;
    DCML_WARNING:
      LevelType := 'WARN';
    DCML_ERROR:
      LevelType := 'ERROR';
  end;
  Self.Lines.Add(Format(MSG_MODEL, [LevelType, messageId, String(messageText), String(fileUrl), line]));
end;

function TConsoleMessageReceiver.WriteWithUrlLineAndColumn(source: PWideChar; level: _DEV_CONSOLE_MESSAGE_LEVEL;
  messageId: SYSINT; messageText, fileUrl: PWideChar; line, column: LongWord): HResult;
const
  MSG_MODEL = '%s'#9'Code %d - %s'#13#10#9'> at %s'#13#10#9'> in line %d, column %d';
var
  LevelType: String;
begin
  Result := S_OK;
  case level of
    DCML_INFORMATIONAL:
      begin
        if not Self.FShowInfo then
          Exit;
        LevelType := 'INFO';
      end;
    DCML_WARNING:
      LevelType := 'WARN';
    DCML_ERROR:
      LevelType := 'ERROR';
  end;
  Self.Lines.Add(Format(MSG_MODEL, [LevelType, messageId, String(messageText), String(fileUrl), line, column]));
end;

end.
