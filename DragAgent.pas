unit DragAgent;

{
  Unit Name   : DragAgent
  Description: 任意の TWinControl に対して、ドラッグ操作による仮想ファイルやテキストなどの
               データ送信機能を追加するための基底クラス群です。
               マウス操作からドラッグを検出し、COMの DoDragDrop によるドラッグ送信を行います。
               データ自体は abstract な `DoDragDataMake` を継承クラスで実装することで柔軟に拡張可能です。

  Features   :
    - 任意の VCL コントロールにドラッグ機能を付与（Attach / Detach）
    - ドラッグ中のマウス座標・状態を管理
    - ドラッグ開始、キャンセル、ターゲット到達イベントに対応
    - IDropSource 実装による正規のドラッグ送信プロトコルに対応
    - ファイル、テキスト、HTMLなどの形式に拡張可能

  Usage      :
    - TDragAgent を継承して `DoDragDataMake` と `DoDragRequest` を実装
    - `Attach(Control)` でコントロールにドラッグ操作を関連付け
    - `OnDragging`, `OnDragCancel`, `OnDragTarget` などのイベントで状態通知を受け取る
    - 派生例: TDragShellFile（仮想ファイルドラッグ用）

  Dependencies:
    - Windows, ActiveX, ShlObj（COM ベースの D&D 処理）

  Notes      :
    - ドラッグ中の制御は IDropSource 経由で行われ、ESC キーやマウスボタンにより中断可能
    - VCL 標準の DragMode とは独立して動作
    - 派生クラスの自由度が高く、複数形式への対応が容易です

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
//  ドラッグ＆ドロップを管理する基礎クラス                                  //
//--------------------------------------------------------------------------//
type

  TDragAgent = class(TWinControl,IDropSource)
  private
    { Private 宣言 }
    FTimer          : TTimer;                   // ダブルクリック後のマウス処理を無効にするタイマー
    FDragStartPos   : TPoint;                   // ドラッグ開始マウス座標
    FDataObject     : IDataObject;              // ドロップ先へ送るデータの管理
    FIsDragging     : Boolean;                  // True :マウスドラッグ中
    FMouseDown      : Boolean;                  // True :マウスクリック中
    FMouseDBClicked : Boolean;                  // True : マウスクリック直後
    FTarget         : TWinControl;              // ターゲット可変にするためTWinControlを使用

    FOnDragTarget   : TDragAgentTarget;         // ドラッグ先ターゲットイベント
    FOnDragCancel   : TNotifyEvent;
    FOnDragging     : TNotifyEvent;

    FOldMouseDown   : TMouseEvent;              // ドラッグ機能を割り当てる前のイベントを保存
    FOldMouseMove   : TMouseMoveEvent;
    FOldMouseUp     : TMouseEvent;

    // データをターゲットに送信
    procedure DataDrop();
    procedure OnTimer(Sender: TObject);
    // マウス降下時に呼ぶ処理 OnMouseDown に直接割り当てても良い
    procedure CtrlMouseDown(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
    // マウス上昇時に呼ぶ処理 OnMouseUp に直接割り当てても良い
    procedure CtrlMouseUp(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
    // マウス移動時に呼ぶ処理 OnMouseMoveに直接割り当てても良い
    procedure CtrlMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
  protected
    // フォルダ名とファイル一覧からターゲットに渡すデータを作成
    procedure DoDragDataMake(const Reset : Boolean);virtual;abstract;
    procedure DoDragTarget(Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
    procedure DoDragCancel();
    procedure DoDragging();
    procedure DoDragRequest();virtual;abstract;
    // ドラッグ中に発生するイベント
    function QueryContinueDrag(fEscapePressed: BOOL; grfKeyState: Longint): HResult; stdcall;
    function GiveFeedback(dwEffect: Longint): HResult; stdcall;
  public
    { Public 宣言 }
    constructor Create(AOwner: TComponent);override;
    destructor Destroy;override;

    // 指定したクラスにドラッグの機能を実装
    procedure Attach(Control: TWinControl);
    // クラスに実装したドラッグの機能を解除
    procedure Detach;
    // ドラッグ操作を強制キャンセル
    procedure CtrlDragCancel(Sender: TObject);

    // ドラッグ元のコントロール
    property Target : TWinControl read FTarget;
    // True : ドラッグ中
    property IsDragging : Boolean read FIsDragging;
    // ドラッグされたときに発生するイベント
    property OnDragTarget : TDragAgentTarget read FOnDragTarget write FOnDragTarget;
    // ドラッグが中断されたときに発生するイベント
    property OnDragCancel : TNotifyEvent read FOnDragCancel write FOnDragCancel;
    // ドラッグが開始されたときに発生するイベント
    property OnDragging : TNotifyEvent read FOnDragging write FOnDragging;
  end;

  // ファイルをドラッグするクラス ※仕組みが古い
type
  TDragShellFile = class(TDragAgent)
  private
    { Private 宣言 }
    FDragFolder     : string;                         // ドロップ先へ送るフォルダ名
    FDragFiles      : TStringList;                    // ドロップ先へ送るファイル一覧(Pathなし)
    FOnDragRequest  : TDragAgentRequestFile;          // ドラッグ時のデータ要求イベント
    // フォルダ名とファイル一覧からターゲットに渡すデータを作成
    function GetFileListDataObject(const Directory: string; Files: TStrings):IDataObject;

    procedure SetDragFiles(ts : TStringList);
  protected
    procedure DoDragDataMake(const Reset : Boolean);override;
    procedure DoDragRequest();override;
    procedure DoDragRequestFiles(FileNames : TStringList);

  public
    { Public 宣言 }
    constructor Create(AOwner: TComponent);override;
    destructor Destroy;override;
    // ドラッグのデータ生成要求イベント
    property OnDragRequest : TDragAgentRequestFile read FOnDragRequest write FOnDragRequest;
  end;

  // ドラッグするためのインターフェイス複数形式に対応
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


  // テキスト関係をドラッグするクラス
type
       TDragTextFormat = set of (dtText, dtHtml);
  TDragText = class(TDragAgent)
  private
    FDragText       : string;                    // ドロップ先へ送るテキスト
    FOnDragRequest  : TDragAgentRequestText;
    FEnabledFormats : TDragTextFormat;
    function GetDataObject(const Text: string):IDataObject;
    // テキストデータからターゲットに渡すデータを作成
    procedure GetTextDataObject(Objects : TDragData;const Text: string);
    procedure GetHtmlDataObject(Objects : TDragData;const Text: string);
  protected
    procedure DoDragDataMake(const Reset : Boolean);override;
    procedure DoDragRequest; override;
    procedure DoDragRequestText(var Text: string);
  public
    constructor Create(AOwner: TComponent);override;
    property EnabledFormats: TDragTextFormat read FEnabledFormats write FEnabledFormats;
    // ドラッグのデータ生成要求イベント
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
  Access := TWinControlAccess(Control);  // ← protectedにアクセスするための型キャスト

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
  //OLEドラッグ＆ドロップ開始
  dwEffect := DROPEFFECT_NONE;
  DoDragDrop(FDataObject, Self, DROPEFFECT_COPY, dwEffect);
end;

procedure TDragAgent.CtrlDragCancel(Sender: TObject);
begin
  FMouseDBClicked := True;
  FTimer.Enabled := True;

  DoDragDataMake(True);                                     // ドラッグデータを初期化
  DoDragCancel();                                     // キャンセルを通知
  FTarget.EndDrag(False);                             // ドラッグ状態を解除
end;


procedure TDragAgent.CtrlMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if Button <> mbLeft then exit;                      // 左クリック降下中でなければ処理しない

  if FMouseDBClicked then exit;                       // マウスクリック直後は処理しない

  FTarget.EndDrag(False);                           // 一度ドラッグを解除
  FMouseDown := True;                               // マウス降下状態を記憶
  DoDragRequest();                                  // D&D用データ作成要求
  DoDragTarget(Button,Shift,X,Y);                   // ドラッグ座標を通知

  FDragStartPos := Point(X, Y);                     // ドラッグ開始座標として記録

end;

procedure TDragAgent.CtrlMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
const
  Threshold = 16;                                          // マウス移動距離
var
  f : Boolean;
begin
  if not FMouseDown then exit;

  f := False;
  if ((Abs(X - FDragStartPos.x) >= Threshold)
    or (Abs(Y - FDragStartPos.y) >= Threshold)) then begin // ボタン降下位置から一定距離離れた場合
    f := True;                                             // ドラッグ中とする
    DoDragging();                                          // ドラッグの開始を通知
  end;
  if not f then exit;                                      // ドラッグ中判定中は処理しない

  if (FMouseDown) and (not FIsDragging) then begin         // マウス降下後の最初のマウス移動の場合
    if FMouseDBClicked then exit;                          // ダブルクリックの可能性がある場合未処理
    FTarget.BeginDrag(True);                               // ドラッグ処理を開始
    FIsDragging := True;                                   // ドラッグ中とする
    DoDragDataMake(False);                                 // データを要求
    DataDrop();                                            // ドラッグ中として振る舞う
  end;
end;

// マウス上昇イベント　※ドラッグ中はドラッグの最中に発生する
procedure TDragAgent.CtrlMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  FMouseDown := False;                                     // マウス降下状態をリセット
  FIsDragging := False;                                    // ドラッグ状態をリセット
  DoDragCancel();                                          // キャンセルを通知
  FTarget.EndDrag(False);                                  // ドラッグ状態を解除
end;

function TDragAgent.QueryContinueDrag(fEscapePressed: BOOL;
  grfKeyState: Integer): HResult;
begin
  if fEscapePressed or
     (grfKeyState and MK_RBUTTON = MK_RBUTTON) then begin  // ドラッグ操作が解除された場合
    result := DRAGDROP_S_CANCEL;                           // アイコンを解除状態に
    DebugLogMsg('DRAGDROP_S_CANCEL');
    FTarget.EndDrag(False);                                // ドラッグを解除
  end
  else if grfKeyState and MK_LBUTTON = 0 then begin        // ドロップ先が決まった場合
    result := DRAGDROP_S_DROP;                             // アイコンをドロップ状態に
    DebugLogMsg('DRAGDROP_S_DROP');
    DoDragCancel();                                        // ドロップの終了処理を通知
    FTarget.EndDrag(False);                                // ドラッグ操作を解除
    FMouseDown := False;                                   // マウス降下状態をリセット
    FIsDragging := False;                                  // ドラッグ状態をリセット
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
    DoDragRequestFiles(FileNames);                    // D&D用データ作成要求
    if FileNames.Count = 0 then exit;                 // データが作成されていない場合は処理しない
    SetDragFiles(FileNames);                          // 指定されたファイル名をドロップデータとする
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
  if Reset then  FDragFiles.Clear;                               // 初期化フラグTrueで初期化

  slst := TStringList.Create;
  try
    if FDragFiles.Count = 0 then exit;
    for i := 0 to FDragFiles.Count-1 do begin                    // ファイルの数だけループ
      slst.Add(FDragFiles[i]);                                   // ファイルを追加
    end;
    FDataObject := GetFileListDataObject(FDragFolder,slst);      //ファイル名からIDataObjectを取得
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
  DoDragRequestText(Text);                          // D&D用データ作成要求
  if Text = '' then exit;                           // データが作成されていない場合は処理しない
  FDragText := Text;                                // 指定されたテキストをドロップデータとする
  FDataObject := GetDataObject(FDragText);                  //ファイル名からIDataObjectを取得
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

  // HTML構造を構築
  HtmlFull := UTF8String(
    '<html><body><!--StartFragment-->' + Text + '<!--EndFragment--></body></html>'
    );

  // オフセットを計算する（ヘッダの長さを見積もり）
  // プレースホルダを用いてあとで置き換える
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

  // 最終HTML全体を組み立て（オフセットを反映）
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

  // FormatEtcを構築
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
  FEnabledFormats := [dtText];  // デフォルトはテキストのみ
end;

procedure TDragText.DoDragDataMake(const Reset: Boolean);
begin
  if Reset then  FDragText := '';                               // 初期化フラグTrueで初期化

  FDataObject := GetDataObject(FDragText);                  //ファイル名からIDataObjectを取得
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
    Exit(E_NOTIMPL);  // GET 以外（SET）は未サポート

  if FFormatList.Count = 0 then
    Exit(S_FALSE);    // 形式が無い場合は列挙不可

  // 一時配列にコピー
  SetLength(FormatArray, FFormatList.Count);
  for i := 0 to FFormatList.Count - 1 do begin
    FormatArray[i] := FFormatList[i];
     //OutputDebugString(PChar(Format('Format[%d]: cfFormat=%d tymed=%d aspect=%d',
     //  [i, FFormatList[i].cfFormat, FFormatList[i].tymed, FFormatList[i].dwAspect])));
  end;

  // Delphi 10 では @FormatArray[0] が型不一致になることがあるので Pointer() を使う
  Result := SHCreateStdEnumFmtEtc(
    Length(FormatArray),
    FormatArray[0], // 明示的に配列の先頭を渡す
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
  // 初期化
  //OutputDebugString(PChar('GetData-------------------'));
  FillChar(Medium, SizeOf(TStgMedium), 0);

  // フォーマットの一致を確認（cfFormatのみ比較でOK）
  for i := 0 to FFormatList.Count - 1 do
  begin
    if FFormatList[i].cfFormat = FormatEtcIn.cfFormat then
    begin
      // 一致したので Medium をコピー
      Medium := FMediumList[i];
      //OutputDebugString(PChar('GetData:'+IntToStr(i)));
      // 呼び出し側に渡すので、自前で複製すべき（簡易対応：そのまま共有）
      // 本来は GlobalAlloc でコピーして渡すと安全

      Result := S_OK;
      Exit;
    end;
  end;

  // フォーマットが見つからなかった
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
    Medium.stm := nil;  // AddRef/Release は COMが管理していると仮定
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

  // 同一 cfFormat がすでにあるかチェック
  for i := 0 to FFormatList.Count - 1 do
  begin
    if FFormatList[i].cfFormat = formatetc.cfFormat then
    begin
      // 上書きする場合、既存Mediumを解放
      m := FMediumList[i];
      ReleaseMedium(m);
      FFormatList[i] := formatetc;
      FMediumList[i] := medium;
      Existing := True;
      Break;
    end;
  end;

  // なければ新規追加
  if not Existing then
  begin
    FFormatList.Add(formatetc);
    FMediumList.Add(medium);
  end;

  // 呼び出し元が解放責任を放棄する場合、自分が保持する
  if not fRelease then
  begin
    // 呼び出し元が保持し続ける → データを自分でコピーすべき（現時点では省略）
    // 実際の用途で問題になる場合は複製処理を追加
  end;

  Result := S_OK;
end;

initialization
  OleInitialize(nil);
finalization
  OleUninitialize;


end.
