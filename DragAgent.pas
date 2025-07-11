unit DragAgent;

{
  Unit Name   : DragAgent
  Description: �C�ӂ� TWinControl �ɑ΂��āA�h���b�O����ɂ�鉼�z�t�@�C����e�L�X�g�Ȃǂ�
               �f�[�^���M�@�\��ǉ����邽�߂̊��N���X�Q�ł��B
               �}�E�X���삩��h���b�O�����o���ACOM�� DoDragDrop �ɂ��h���b�O���M���s���܂��B
               �f�[�^���̂� abstract �� `DoDragDataMake` ���p���N���X�Ŏ������邱�Ƃŏ_��Ɋg���\�ł��B

  Features   :
    - �C�ӂ� VCL �R���g���[���Ƀh���b�O�@�\��t�^�iAttach / Detach�j
    - �h���b�O���̃}�E�X���W�E��Ԃ��Ǘ�
    - �h���b�O�J�n�A�L�����Z���A�^�[�Q�b�g���B�C�x���g�ɑΉ�
    - IDropSource �����ɂ�鐳�K�̃h���b�O���M�v���g�R���ɑΉ�
    - �t�@�C���A�e�L�X�g�AHTML�Ȃǂ̌`���Ɋg���\

  Usage      :
    - TDragAgent ���p������ `DoDragDataMake` �� `DoDragRequest` ������
    - `Attach(Control)` �ŃR���g���[���Ƀh���b�O������֘A�t��
    - `OnDragging`, `OnDragCancel`, `OnDragTarget` �Ȃǂ̃C�x���g�ŏ�Ԓʒm���󂯎��
    - �h����: TDragShellFile�i���z�t�@�C���h���b�O�p�j

  Dependencies:
    - Windows, ActiveX, ShlObj�iCOM �x�[�X�� D&D �����j

  Notes      :
    - �h���b�O���̐���� IDropSource �o�R�ōs���AESC �L�[��}�E�X�{�^���ɂ�蒆�f�\
    - VCL �W���� DragMode �Ƃ͓Ɨ����ē���
    - �h���N���X�̎��R�x�������A�����`���ւ̑Ή����e�Ղł�

  Author     : vramwiz
  Created    : 2025-07-10
  Updated    : 2025-07-10
}

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, ActiveX, ShlObj, ComObj,
  Vcl.StdCtrls,System.Types,Vcl.ExtCtrls,System.Generics.Collections;

type TDragAgentRequestFile = procedure (Sender : TObject;FileNames : TStringList) of object;
type TDragAgentRequestText = procedure (Sender : TObject;var Text : string) of object;
type TDragAgentTarget = procedure (Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer) of object;


//--------------------------------------------------------------------------//
//  �h���b�O���h���b�v���Ǘ������b�N���X                                  //
//--------------------------------------------------------------------------//
type

  TDragAgent = class(TWinControl,IDropSource)
  private
    { Private �錾 }
    FTimer          : TTimer;                   // �_�u���N���b�N��̃}�E�X�����𖳌��ɂ���^�C�}�[
    FDragStartPos   : TPoint;                   // �h���b�O�J�n�}�E�X���W
    FDataObject     : IDataObject;              // �h���b�v��֑���f�[�^�̊Ǘ�
    FIsDragging     : Boolean;                  // True :�}�E�X�h���b�O��
    FMouseDown      : Boolean;                  // True :�}�E�X�N���b�N��
    FMouseDBClicked : Boolean;                  // True : �}�E�X�N���b�N����
    FTarget         : TWinControl;              // �^�[�Q�b�g�ςɂ��邽��TWinControl���g�p

    FOnDragTarget   : TDragAgentTarget;         // �h���b�O��^�[�Q�b�g�C�x���g
    FOnDragCancel   : TNotifyEvent;
    FOnDragging     : TNotifyEvent;

    FOldMouseDown   : TMouseEvent;              // �h���b�O�@�\�����蓖�Ă�O�̃C�x���g��ۑ�
    FOldMouseMove   : TMouseMoveEvent;
    FOldMouseUp     : TMouseEvent;

    // �f�[�^���^�[�Q�b�g�ɑ��M
    procedure DataDrop();
    procedure OnTimer(Sender: TObject);
    // �}�E�X�~�����ɌĂԏ��� OnMouseDown �ɒ��ڊ��蓖�ĂĂ��ǂ�
    procedure CtrlMouseDown(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
    // �}�E�X�㏸���ɌĂԏ��� OnMouseUp �ɒ��ڊ��蓖�ĂĂ��ǂ�
    procedure CtrlMouseUp(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
    // �}�E�X�ړ����ɌĂԏ��� OnMouseMove�ɒ��ڊ��蓖�ĂĂ��ǂ�
    procedure CtrlMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
  protected
    // �t�H���_���ƃt�@�C���ꗗ����^�[�Q�b�g�ɓn���f�[�^���쐬
    procedure DoDragDataMake(const Reset : Boolean);virtual;abstract;
    procedure DoDragTarget(Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
    procedure DoDragCancel();
    procedure DoDragging();
    procedure DoDragRequest();virtual;abstract;
    // �h���b�O���ɔ�������C�x���g
    function QueryContinueDrag(fEscapePressed: BOOL; grfKeyState: Longint): HResult; stdcall;
    function GiveFeedback(dwEffect: Longint): HResult; stdcall;
  public
    { Public �錾 }
    constructor Create(AOwner: TComponent);override;
    destructor Destroy;override;

    // �w�肵���N���X�Ƀh���b�O�̋@�\������
    procedure Attach(Control: TWinControl);
    // �N���X�Ɏ��������h���b�O�̋@�\������
    procedure Detach;
    // �h���b�O����������L�����Z��
    procedure CtrlDragCancel(Sender: TObject);

    // �h���b�O���̃R���g���[��
    property Target : TWinControl read FTarget;
    // True : �h���b�O��
    property IsDragging : Boolean read FIsDragging;
    // �h���b�O���ꂽ�Ƃ��ɔ�������C�x���g
    property OnDragTarget : TDragAgentTarget read FOnDragTarget write FOnDragTarget;
    // �h���b�O�����f���ꂽ�Ƃ��ɔ�������C�x���g
    property OnDragCancel : TNotifyEvent read FOnDragCancel write FOnDragCancel;
    // �h���b�O���J�n���ꂽ�Ƃ��ɔ�������C�x���g
    property OnDragging : TNotifyEvent read FOnDragging write FOnDragging;
  end;

  // �t�@�C�����h���b�O����N���X ���d�g�݂��Â�
type
  TDragShellFile = class(TDragAgent)
  private
    { Private �錾 }
    FDragFolder     : string;                         // �h���b�v��֑���t�H���_��
    FDragFiles      : TStringList;                    // �h���b�v��֑���t�@�C���ꗗ(Path�Ȃ�)
    FOnDragRequest  : TDragAgentRequestFile;          // �h���b�O���̃f�[�^�v���C�x���g
    // �t�H���_���ƃt�@�C���ꗗ����^�[�Q�b�g�ɓn���f�[�^���쐬
    function GetFileListDataObject(const Directory: string; Files: TStrings):IDataObject;

    procedure SetDragFiles(ts : TStringList);
  protected
    procedure DoDragDataMake(const Reset : Boolean);override;
    procedure DoDragRequest();override;
    procedure DoDragRequestFiles(FileNames : TStringList);

  public
    { Public �錾 }
    constructor Create(AOwner: TComponent);override;
    destructor Destroy;override;
    // �h���b�O�̃f�[�^�����v���C�x���g
    property OnDragRequest : TDragAgentRequestFile read FOnDragRequest write FOnDragRequest;
  end;

  // �h���b�O���邽�߂̃C���^�[�t�F�C�X�����`���ɑΉ�
type
  TDragData = class(TInterfacedObject, IDataObject)
  private
    procedure ReleaseMedium(var Medium: TStgMedium);
  protected
    FFormatList: TList<TFormatEtc>;
    FMediumList: TList<TStgMedium>;
  public
    constructor Create;reintroduce;
    destructor Destroy;override;


    function SetData(const formatetc: TFormatEtc; var medium: TStgMedium; fRelease: BOOL): HResult; stdcall;
    function GetData(const FormatEtcIn: TFormatEtc; out Medium: TStgMedium): HResult; stdcall;
    function GetDataHere(const FormatEtc: TFormatEtc; out Medium: TStgMedium): HResult; stdcall;
    function GetCanonicalFormatEtc(const FormatEtc: TFormatEtc; out FormatEtcOut: TFormatEtc): HResult; stdcall;
    function DAdvise(const FormatEtc: TFormatEtc; advf: Longint;
      const AdvSink: IAdviseSink; out dwConnection: Longint): HResult; stdcall;
    function DUnadvise(dwConnection: Longint): HResult; stdcall;
    function EnumDAdvise(out EnumAdvise: IEnumStatData): HResult; stdcall;
    function QueryGetData(const FormatEtc: TFormatEtc): HResult; stdcall;
    function EnumFormatEtc(dwDirection: Integer; out EnumFormatEtc: IEnumFormatEtc): HResult; stdcall;
  end;


  // �e�L�X�g�֌W���h���b�O����N���X
type
       TDragTextFormat = set of (dtText, dtHtml);
  TDragText = class(TDragAgent)
  private
    FDragText       : string;                    // �h���b�v��֑���e�L�X�g
    FOnDragRequest  : TDragAgentRequestText;
    FEnabledFormats : TDragTextFormat;
    function GetDataObject(const Text: string):IDataObject;
    // �e�L�X�g�f�[�^����^�[�Q�b�g�ɓn���f�[�^���쐬
    procedure GetTextDataObject(Objects : TDragData;const Text: string);
    procedure GetHtmlDataObject(Objects : TDragData;const Text: string);
  protected
    procedure DoDragDataMake(const Reset : Boolean);override;
    procedure DoDragRequest; override;
    procedure DoDragRequestText(var Text: string);
  public
    constructor Create(AOwner: TComponent);override;
    property EnabledFormats: TDragTextFormat read FEnabledFormats write FEnabledFormats;
    // �h���b�O�̃f�[�^�����v���C�x���g
    property OnDragRequest: TDragAgentRequestText read FOnDragRequest write FOnDragRequest;
  end;


implementation

procedure DebugLogMsg(const Msg: string);
begin
  {$IFDEF DEBUG}
  OutputDebugString(PChar(Msg));
  {$ENDIF}
end;



{ TDragFile }

constructor TDragAgent.Create(AOwner: TComponent);
begin
  inherited;
  FTimer := TTimer.Create(nil);
  FTimer.Interval := 1000;
  FTimer.Enabled := False;
  FTimer.OnTimer := OnTimer;
end;

destructor TDragAgent.Destroy;
begin
  FTimer.Free;
  inherited;
end;

type
  TWinControlAccess = class(TWinControl);

procedure TDragAgent.Attach(Control: TWinControl);
var
  Access: TWinControlAccess;
begin
  if Assigned(FTarget) then
    Detach;

  FTarget := Control;
  Access := TWinControlAccess(Control);  // �� protected�ɃA�N�Z�X���邽�߂̌^�L���X�g

  FOldMouseDown := Access.OnMouseDown;
  FOldMouseMove := Access.OnMouseMove;
  FOldMouseUp   := Access.OnMouseUp;

  Access.OnMouseDown := CtrlMouseDown;
  Access.OnMouseMove := CtrlMouseMove;
  Access.OnMouseUp   := CtrlMouseUp;
end;

procedure TDragAgent.Detach;
var
  Access: TWinControlAccess;
begin
  if not Assigned(FTarget) then Exit;

  Access := TWinControlAccess(FTarget);
  Access.OnMouseDown := FOldMouseDown;
  Access.OnMouseMove := FOldMouseMove;
  Access.OnMouseUp   := FOldMouseUp;

  FTarget := nil;
end;

function TDragAgent.GiveFeedback(dwEffect: Integer): HResult;
begin
  result := DRAGDROP_S_USEDEFAULTCURSORS;
end;

procedure TDragAgent.OnTimer(Sender: TObject);
begin
  FMouseDBClicked := False;
end;

procedure TDragAgent.DataDrop;
var
  dwEffect : Integer;
begin
  //OLE�h���b�O���h���b�v�J�n
  dwEffect := DROPEFFECT_NONE;
  DoDragDrop(FDataObject, Self, DROPEFFECT_COPY, dwEffect);
end;

procedure TDragAgent.CtrlDragCancel(Sender: TObject);
begin
  FMouseDBClicked := True;
  FTimer.Enabled := True;

  DoDragDataMake(True);                                     // �h���b�O�f�[�^��������
  DoDragCancel();                                     // �L�����Z����ʒm
  FTarget.EndDrag(False);                             // �h���b�O��Ԃ�����
end;


procedure TDragAgent.CtrlMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button <> mbLeft then exit;                      // ���N���b�N�~�����łȂ���Ώ������Ȃ�

  if FMouseDBClicked then exit;                       // �}�E�X�N���b�N����͏������Ȃ�

  FTarget.EndDrag(False);                           // ��x�h���b�O������
  FMouseDown := True;                               // �}�E�X�~����Ԃ��L��
  DoDragRequest();                                  // D&D�p�f�[�^�쐬�v��
  DoDragTarget(Button,Shift,X,Y);                   // �h���b�O���W��ʒm

  FDragStartPos := Point(X, Y);                     // �h���b�O�J�n���W�Ƃ��ċL�^

end;

procedure TDragAgent.CtrlMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
const
  Threshold = 16;                                          // �}�E�X�ړ�����
var
  f : Boolean;
begin
  if not FMouseDown then exit;

  f := False;
  if ((Abs(X - FDragStartPos.x) >= Threshold)
    or (Abs(Y - FDragStartPos.y) >= Threshold)) then begin // �{�^���~���ʒu�����苗�����ꂽ�ꍇ
    f := True;                                             // �h���b�O���Ƃ���
    DoDragging();                                          // �h���b�O�̊J�n��ʒm
  end;
  if not f then exit;                                      // �h���b�O�����蒆�͏������Ȃ�

  if (FMouseDown) and (not FIsDragging) then begin         // �}�E�X�~����̍ŏ��̃}�E�X�ړ��̏ꍇ
    if FMouseDBClicked then exit;                          // �_�u���N���b�N�̉\��������ꍇ������
    FTarget.BeginDrag(True);                               // �h���b�O�������J�n
    FIsDragging := True;                                   // �h���b�O���Ƃ���
    DoDragDataMake(False);                                 // �f�[�^��v��
    DataDrop();                                            // �h���b�O���Ƃ��ĐU�镑��
  end;
end;

// �}�E�X�㏸�C�x���g�@���h���b�O���̓h���b�O�̍Œ��ɔ�������
procedure TDragAgent.CtrlMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  FMouseDown := False;                                     // �}�E�X�~����Ԃ����Z�b�g
  FIsDragging := False;                                    // �h���b�O��Ԃ����Z�b�g
  DoDragCancel();                                          // �L�����Z����ʒm
  FTarget.EndDrag(False);                                  // �h���b�O��Ԃ�����
end;

function TDragAgent.QueryContinueDrag(fEscapePressed: BOOL;
  grfKeyState: Integer): HResult;
begin
  if fEscapePressed or
     (grfKeyState and MK_RBUTTON = MK_RBUTTON) then begin  // �h���b�O���삪�������ꂽ�ꍇ
    result := DRAGDROP_S_CANCEL;                           // �A�C�R����������Ԃ�
    DebugLogMsg('DRAGDROP_S_CANCEL');
    FTarget.EndDrag(False);                                // �h���b�O������
  end
  else if grfKeyState and MK_LBUTTON = 0 then begin        // �h���b�v�悪���܂����ꍇ
    result := DRAGDROP_S_DROP;                             // �A�C�R�����h���b�v��Ԃ�
    DebugLogMsg('DRAGDROP_S_DROP');
    DoDragCancel();                                        // �h���b�v�̏I��������ʒm
    FTarget.EndDrag(False);                                // �h���b�O���������
    FMouseDown := False;                                   // �}�E�X�~����Ԃ����Z�b�g
    FIsDragging := False;                                  // �h���b�O��Ԃ����Z�b�g
  end else begin
    result := S_OK;
  end;

end;


procedure TDragAgent.DoDragCancel;
begin
  DebugLogMsg('DO_DRAG_CANCEL');
  if Assigned(FOnDragCancel) then begin
    FOnDragCancel(Self);
  end;
end;

procedure TDragAgent.DoDragging;
begin
  if Assigned(FOnDragging) then begin
    FOnDragging(Self);
  end;
end;

procedure TDragAgent.DoDragTarget(Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
begin
  if Assigned(FOnDragTarget) then begin
    FOnDragTarget(Self,Button,Shift,X,Y);
  end;
end;


{ TDragFile }

constructor TDragShellFile.Create(AOwner: TComponent);
begin
  inherited;
  FDragFiles := TStringList.Create;
end;

destructor TDragShellFile.Destroy;
begin
  FDragFiles.Free;
  inherited;
end;

procedure TDragShellFile.DoDragRequest;
var
  FileNames : TStringList;
begin
  FileNames := TStringList.Create;
  try
    DoDragRequestFiles(FileNames);                    // D&D�p�f�[�^�쐬�v��
    if FileNames.Count = 0 then exit;                 // �f�[�^���쐬����Ă��Ȃ��ꍇ�͏������Ȃ�
    SetDragFiles(FileNames);                          // �w�肳�ꂽ�t�@�C�������h���b�v�f�[�^�Ƃ���
  finally
    FileNames.Free;
  end;
end;

procedure TDragShellFile.DoDragRequestFiles(FileNames: TStringList);
begin
  DebugLogMsg('DoDragRequestFiles');
  if Assigned(FOnDragRequest) then FOnDragRequest(Self,FileNames);
end;

function TDragShellFile.GetFileListDataObject(const Directory: string; Files: TStrings): IDataObject;
type
  PArrayOfPItemIDList = ^TArrayOfPItemIDList;
  TArrayOfPItemIDList = array[0..0] of PItemIDList;
var
  Malloc       : IMalloc;
  Root         : IShellFolder;
  FolderPidl   : PItemIDList;
  Folder       : IShellFolder;
  chEaten      : ULONG;
  dwAttributes : ULONG;
  FileCount    : Integer;
  p            : array of PItemIDList;
  i            : Integer;
begin
  Result := nil;
  if Files.Count = 0 then Exit;
  OleCheck(SHGetMalloc(Malloc));
  OleCheck(SHGetDesktopFolder(Root));
  OleCheck(Root.ParseDisplayName(0, nil, PWideChar(WideString(Directory)),
    chEaten, FolderPidl, dwAttributes));
  try
    OleCheck(Root.BindToObject(FolderPidl, nil, IShellFolder,
      Pointer(Folder)));
    FileCount := Files.Count;
    SetLength(p,FileCount);
    for i := 0 to FileCount - 1 do begin
      OleCheck(Folder.ParseDisplayName(0, nil,
        PWideChar(WideString(Files[i])), chEaten, p[i], dwAttributes));
    end;
    OleCheck(Folder.GetUIObjectOf(0, FileCount, p[0], IDataObject, nil,
      Pointer(Result)));
  finally
    Malloc.Free(FolderPidl);
  end;
end;

procedure TDragShellFile.DoDragDataMake(const Reset : Boolean);
var
  slst : TStringList;
  i : Integer;
begin
  if Reset then  FDragFiles.Clear;                               // �������t���OTrue�ŏ�����

  slst := TStringList.Create;
  try
    if FDragFiles.Count = 0 then exit;
    for i := 0 to FDragFiles.Count-1 do begin                    // �t�@�C���̐��������[�v
      slst.Add(FDragFiles[i]);                                   // �t�@�C����ǉ�
    end;
    FDataObject := GetFileListDataObject(FDragFolder,slst);      //�t�@�C��������IDataObject���擾
  finally
    slst.Free;
  end;
end;

procedure TDragShellFile.SetDragFiles(ts: TStringList);
var
  i: Integer;
begin
  if ts.Count = 0 then exit;
  FDragFolder := ExtractFilePath(ts[0]);
  FDragFiles.Clear;
  for i := 0 to ts.Count-1 do begin
    FDragFiles.Add(ExtractFileName(ts[i]));
  end;
end;


{ TDragText }

procedure TDragText.DoDragRequest;
var
  Text : string;
begin
  DoDragRequestText(Text);                          // D&D�p�f�[�^�쐬�v��
  if Text = '' then exit;                           // �f�[�^���쐬����Ă��Ȃ��ꍇ�͏������Ȃ�
  FDragText := Text;                                // �w�肳�ꂽ�e�L�X�g���h���b�v�f�[�^�Ƃ���
  FDataObject := GetDataObject(FDragText);                  //�t�@�C��������IDataObject���擾
end;

procedure TDragText.DoDragRequestText(var Text: string);
begin
  if Assigned(FOnDragRequest) then FOnDragRequest(Self,Text);
end;


function MakeFormatEtc(cfFormat: Word; tymed: Integer): TFormatEtc;
begin
  Result.cfFormat := cfFormat;
  Result.ptd := nil;
  Result.dwAspect := DVASPECT_CONTENT;
  Result.lindex := -1;
  Result.tymed := tymed;
end;


function TDragText.GetDataObject(const Text: string): IDataObject;
var
  Objects : TDragData;
begin
  Objects := TDragData.Create;
  if dtText in FEnabledFormats then GetTextDataObject(Objects,Text);
  if dtHtml in FEnabledFormats then GetHtmlDataObject(Objects,Text);
  Result := Objects;
end;

procedure TDragText.GetHtmlDataObject(Objects: TDragData; const Text: string);
var
  Medium: TStgMedium;
  FormatEtc: TFormatEtc;
  MemHandle: HGLOBAL;
  HtmlFull, HtmlUtf8: UTF8String;
  P: Pointer;
  Size: Integer;
  CF_HTML_FORMAT: UINT;
  Header: string;
  StartHTML, EndHTML, StartFragment, EndFragment: Integer;
begin
  if Text = '' then Exit;

  // HTML�\�����\�z
  HtmlFull := UTF8String(
    '<html><body><!--StartFragment-->' + Text + '<!--EndFragment--></body></html>'
    );

  // �I�t�Z�b�g���v�Z����i�w�b�_�̒��������ς���j
  // �v���[�X�z���_��p���Ă��ƂŒu��������
  Header :=
    'Version:1.0'#13#10 +
    'StartHTML:00000000'#13#10 +
    'EndHTML:00000000'#13#10 +
    'StartFragment:00000000'#13#10 +
    'EndFragment:00000000'#13#10;

  StartHTML := Length(Header);
  StartFragment := Pos('<!--StartFragment-->', string(HtmlFull)) - 1;
  EndFragment := Pos('<!--EndFragment-->', string(HtmlFull)) + Length('<!--EndFragment-->') - 1;
  EndHTML := StartHTML + Length(HtmlFull);

  // �ŏIHTML�S�̂�g�ݗ��āi�I�t�Z�b�g�𔽉f�j
  Header :=
    Format('Version:1.0'#13#10 +
           'StartHTML:%08d'#13#10 +
           'EndHTML:%08d'#13#10 +
           'StartFragment:%08d'#13#10 +
           'EndFragment:%08d'#13#10,
      [StartHTML, EndHTML, StartHTML + StartFragment, StartHTML + EndFragment]);

  HtmlUtf8 := UTF8String(Header + string(HtmlFull));
  Size := Length(HtmlUtf8) + 1;

  MemHandle := GlobalAlloc(GMEM_MOVEABLE or GMEM_ZEROINIT, Size);
  if MemHandle = 0 then Exit;

  P := GlobalLock(MemHandle);
  try
    Move(PAnsiChar(HtmlUtf8)^, P^, Size);
  finally
    GlobalUnlock(MemHandle);
  end;

  // FormatEtc���\�z
  CF_HTML_FORMAT := RegisterClipboardFormat('HTML Format');
  FormatEtc := MakeFormatEtc(CF_HTML_FORMAT, TYMED_HGLOBAL);

  Medium.tymed := TYMED_HGLOBAL;
  Medium.hGlobal := MemHandle;
  Medium.unkForRelease := nil;

  Objects.SetData(FormatEtc, Medium, True);
end;

procedure TDragText.GetTextDataObject(Objects : TDragData;const Text: string);
var
  Medium: TStgMedium;
  FormatEtc: TFormatEtc;
  MemHandle: HGLOBAL;
  WideText: PWideChar;
  Size: Integer;
begin
  if Text = '' then Exit;

  Size := (Length(Text) + 1) * SizeOf(WideChar);
  MemHandle := GlobalAlloc(GMEM_MOVEABLE or GMEM_ZEROINIT, Size);
  if MemHandle = 0 then Exit;

  WideText := GlobalLock(MemHandle);
  try
    Move(PChar(Text)^, WideText^, Size);
  finally
    GlobalUnlock(MemHandle);
  end;

  FormatEtc := MakeFormatEtc(CF_UNICODETEXT, TYMED_HGLOBAL);

  Medium.tymed := TYMED_HGLOBAL;
  Medium.hGlobal := MemHandle;
  Medium.unkForRelease := nil;

  Objects.SetData(FormatEtc, Medium, True);
end;



constructor TDragText.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FEnabledFormats := [dtText];  // �f�t�H���g�̓e�L�X�g�̂�
end;

procedure TDragText.DoDragDataMake(const Reset: Boolean);
begin
  if Reset then  FDragText := '';                               // �������t���OTrue�ŏ�����

  FDataObject := GetDataObject(FDragText);                  //�t�@�C��������IDataObject���擾
end;



{ TDragData }

constructor TDragData.Create;
begin
  inherited;
  FFormatList := TList<TFormatEtc>.Create;
  FMediumList := TList<TStgMedium>.Create;
end;

destructor TDragData.Destroy;
var
  i: Integer;
  m : TStgMedium;
begin
  for i := 0 to FMediumList.Count - 1 do
    m := FMediumList[i];
    ReleaseMedium(m);

  FFormatList.Free;
  FMediumList.Free;

  inherited Destroy;
end;

function TDragData.DAdvise(const FormatEtc: TFormatEtc; advf: Integer;
  const AdvSink: IAdviseSink; out dwConnection: Integer): HResult;
begin
  Result := OLE_E_ADVISENOTSUPPORTED;
end;

function TDragData.DUnadvise(dwConnection: Integer): HResult;
begin
  Result := OLE_E_ADVISENOTSUPPORTED;
end;

function TDragData.EnumDAdvise(out EnumAdvise: IEnumStatData): HResult;
begin
  Result := OLE_E_ADVISENOTSUPPORTED;
end;

function TDragData.EnumFormatEtc(dwDirection: Integer;
  out EnumFormatEtc: IEnumFormatEtc): HResult;
var
  FormatArray: array of TFormatEtc;
  i: Integer;
begin
  EnumFormatEtc := nil;

  if dwDirection <> DATADIR_GET then
    Exit(E_NOTIMPL);  // GET �ȊO�iSET�j�͖��T�|�[�g

  if FFormatList.Count = 0 then
    Exit(S_FALSE);    // �`���������ꍇ�͗񋓕s��

  // �ꎞ�z��ɃR�s�[
  SetLength(FormatArray, FFormatList.Count);
  for i := 0 to FFormatList.Count - 1 do begin
    FormatArray[i] := FFormatList[i];
     //OutputDebugString(PChar(Format('Format[%d]: cfFormat=%d tymed=%d aspect=%d',
     //  [i, FFormatList[i].cfFormat, FFormatList[i].tymed, FFormatList[i].dwAspect])));
  end;

  // Delphi 10 �ł� @FormatArray[0] ���^�s��v�ɂȂ邱�Ƃ�����̂� Pointer() ���g��
  Result := SHCreateStdEnumFmtEtc(
    Length(FormatArray),
    FormatArray[0], // �����I�ɔz��̐擪��n��
    EnumFormatEtc
  );
end;

function TDragData.GetCanonicalFormatEtc(const FormatEtc: TFormatEtc;
  out FormatEtcOut: TFormatEtc): HResult;
begin
  Result := E_NOTIMPL;
end;

function TDragData.GetData(const FormatEtcIn: TFormatEtc;
  out Medium: TStgMedium): HResult;
var
  i: Integer;
begin
  // ������
  //OutputDebugString(PChar('GetData-------------------'));
  FillChar(Medium, SizeOf(TStgMedium), 0);

  // �t�H�[�}�b�g�̈�v���m�F�icfFormat�̂ݔ�r��OK�j
  for i := 0 to FFormatList.Count - 1 do
  begin
    if FFormatList[i].cfFormat = FormatEtcIn.cfFormat then
    begin
      // ��v�����̂� Medium ���R�s�[
      Medium := FMediumList[i];
      //OutputDebugString(PChar('GetData:'+IntToStr(i)));
      // �Ăяo�����ɓn���̂ŁA���O�ŕ������ׂ��i�ȈՑΉ��F���̂܂܋��L�j
      // �{���� GlobalAlloc �ŃR�s�[���ēn���ƈ��S

      Result := S_OK;
      Exit;
    end;
  end;

  // �t�H�[�}�b�g��������Ȃ�����
  Result := DV_E_FORMATETC;
end;

function TDragData.GetDataHere(const FormatEtc: TFormatEtc;
  out Medium: TStgMedium): HResult;
begin
  Result := E_NOTIMPL;
end;

function TDragData.QueryGetData(const FormatEtc: TFormatEtc): HResult;
var
  i: Integer;
begin
  for i := 0 to FFormatList.Count - 1 do
  begin
    if FFormatList[i].cfFormat = FormatEtc.cfFormat then
    begin
      Result := S_OK;
      Exit;
    end;
  end;

  Result := DV_E_FORMATETC;
end;

procedure TDragData.ReleaseMedium(var Medium: TStgMedium);
begin
  if Medium.tymed = TYMED_HGLOBAL then
    GlobalFree(Medium.hGlobal)
  else if Medium.tymed = TYMED_ISTREAM then
    Medium.stm := nil;  // AddRef/Release �� COM���Ǘ����Ă���Ɖ���
end;


function TDragData.SetData(const formatetc: TFormatEtc; var medium: TStgMedium;
  fRelease: BOOL): HResult;
var
  i: Integer;
  Existing: Boolean;
  m : TStgMedium;
begin
  Result := E_FAIL;
  Existing := False;

  // ���� cfFormat �����łɂ��邩�`�F�b�N
  for i := 0 to FFormatList.Count - 1 do
  begin
    if FFormatList[i].cfFormat = formatetc.cfFormat then
    begin
      // �㏑������ꍇ�A����Medium�����
      m := FMediumList[i];
      ReleaseMedium(m);
      FFormatList[i] := formatetc;
      FMediumList[i] := medium;
      Existing := True;
      Break;
    end;
  end;

  // �Ȃ���ΐV�K�ǉ�
  if not Existing then
  begin
    FFormatList.Add(formatetc);
    FMediumList.Add(medium);
  end;

  // �Ăяo����������ӔC���������ꍇ�A�������ێ�����
  if not fRelease then
  begin
    // �Ăяo�������ێ��������� �� �f�[�^�������ŃR�s�[���ׂ��i�����_�ł͏ȗ��j
    // ���ۂ̗p�r�Ŗ��ɂȂ�ꍇ�͕���������ǉ�
  end;

  Result := S_OK;
end;

initialization
  OleInitialize(nil);
finalization
  OleUninitialize;


end.
