unit MainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,DragAgent, Vcl.StdCtrls, Vcl.ExtCtrls;

type
  TFormMain = class(TForm)
    Panel1: TPanel;
    ListBox1: TListBox;
    Panel2: TPanel;
    ListBox2: TListBox;
    Panel3: TPanel;
    ListBox3: TListBox;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private 宣言 }
    FDragText : TDragText;
    FDragHtml : TDragText;
    FDragFile : TDragShellFile;
    procedure OnDragRequestFile(Sender : TObject;FileNames : TStringList);
    procedure OnDragRequestText(Sender : TObject;var Text : string);
    procedure OnDragRequestHtml(Sender : TObject;var Text : string);
  public
    { Public 宣言 }
  end;

var
  FormMain: TFormMain;

implementation

uses
  System.IOUtils, System.Types;

{$R *.dfm}

procedure TFormMain.FormCreate(Sender: TObject);
begin
  FDragText := TDragText.Create(Self);
  FDragText.OnDragRequest := OnDragRequestText;
  FDragText.EnabledFormats := [dtText];
  FDragText.Attach(ListBox1);

  FDragHtml := TDragText.Create(Self);
  FDragHtml.OnDragRequest := OnDragRequestHtml;
  FDragHtml.EnabledFormats := [dtHtml];
  FDragHtml.Attach(ListBox2);

  FDragFile := TDragShellFile.Create(Self);
  FDragFile.OnDragRequest := OnDragRequestFile;
  FDragFile.Attach(ListBox3);
end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin
  FDragFile.Free;
  FDragHtml.Free;
  FDragText.Free;
//
end;

procedure TFormMain.FormShow(Sender: TObject);
var
  PicPath: string;
  Files: TStringDynArray;
  I: Integer;
begin
  // ListBox1: 通常テキスト
  ListBox1.Items.BeginUpdate;
  try
    ListBox1.Items.Clear;
    ListBox1.Items.Add('これはテキストの1行目です。');
    ListBox1.Items.Add('2行目：コピーしてみてください。');
    ListBox1.Items.Add('最後の行です。');
  finally
    ListBox1.Items.EndUpdate;
  end;

  // ListBox2: HTMLっぽいテキスト
  ListBox2.Items.BeginUpdate;
  try
    ListBox2.Items.Clear;
    ListBox2.Items.Add('<h1>サンプル見出し</h1>');
    ListBox2.Items.Add('<p>これは段落の例です。</p>');
    ListBox2.Items.Add('<ul><li>リスト項目</li></ul>');
  finally
    ListBox2.Items.EndUpdate;
  end;

  // ListBox3: ピクチャフォルダ内のファイル
  PicPath := TPath.GetPicturesPath;
  ListBox3.Items.BeginUpdate;
  try
    ListBox3.Items.Clear;
    if DirectoryExists(PicPath) then
    begin
      Files := TDirectory.GetFiles(PicPath, '*.*');
      for I := 0 to High(Files) do
        ListBox3.Items.Add(ExtractFileName(Files[I]));  // ← 表示は短く
    end
    else
      ListBox3.Items.Add('(ピクチャフォルダが見つかりません)');
  finally
    ListBox3.Items.EndUpdate;
  end;
end;

procedure TFormMain.OnDragRequestText(Sender: TObject; var Text: string);
begin
  Text := ListBox1.Items[ListBox1.ItemIndex];
end;

procedure TFormMain.OnDragRequestHtml(Sender: TObject; var Text: string);
begin
  Text := ListBox2.Items[ListBox2.ItemIndex];
end;

procedure TFormMain.OnDragRequestFile(Sender: TObject; FileNames: TStringList);
var
  FullPath: string;
begin
  FullPath := TPath.Combine(TPath.GetPicturesPath, ListBox3.Items[ListBox3.ItemIndex]);
  FileNames.Add(FullPath);
end;

end.
