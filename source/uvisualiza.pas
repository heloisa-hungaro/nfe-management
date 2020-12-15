unit uvisualiza;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, sqldb, db, FileUtil, Forms, Controls, Graphics, Dialogs,
  DBGrids, StdCtrls, ExtCtrls;

type

  { TFVisualiza }

  TFVisualiza = class(TForm)
    BtnCancelar: TButton;
    D: TDataSource;
    DNotas: TDataSource;
    GridN: TDBGrid;
    GridP: TDBGrid;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    MemoE: TMemo;
    MemoD: TMemo;
    QNotas: TSQLQuery;
    Q: TSQLQuery;
    procedure BtnCancelarClick(Sender: TObject);
    procedure DNotasDataChange(Sender: TObject; Field: TField);
    procedure FormShow(Sender: TObject);
    procedure GridNCellClick(Column: TColumn);
    procedure GridNCellClick(Sender: TObject);
  private
    { private declarations }
  public
    user: integer;
    { public declarations }
  end;

var
  FVisualiza: TFVisualiza;

implementation

{$R *.lfm}

{ TFVisualiza }

procedure TFVisualiza.BtnCancelarClick(Sender: TObject);
begin
  Close;
end;

procedure TFVisualiza.DNotasDataChange(Sender: TObject; Field: TField);
begin
     MemoE.Clear;
     MemoE.Lines.Add(QNotas.FieldByName('emi_rzo').AsString+' ('+QNotas.FieldByName('emi_fan').AsString+')');
     MemoE.Lines.Add('Documento: '+  QNotas.FieldByName('emi_doc').AsString);
     MemoE.Lines.Add('Endereço: '+  QNotas.FieldByName('emi_end').AsString);
     MemoE.Lines.Add('Cidade: '+  QNotas.FieldByName('emi_cid').AsString);
     MemoE.Lines.Add('Telefone: '+  QNotas.FieldByName('emi_tel').AsString);
     MemoE.Lines.Add('E-mail: '+  QNotas.FieldByName('emi_eml').AsString);

     MemoD.Clear;
     MemoD.Lines.Add(QNotas.FieldByName('des_nom').AsString);
     MemoD.Lines.Add('Documento: '+  QNotas.FieldByName('des_doc').AsString);
     MemoD.Lines.Add('Endereço: '+  QNotas.FieldByName('des_end').AsString);
     MemoD.Lines.Add('Cidade: '+  QNotas.FieldByName('des_cid').AsString);
     MemoD.Lines.Add('Telefone: '+  QNotas.FieldByName('des_tel').AsString);
     MemoD.Lines.Add('E-mail: '+  QNotas.FieldByName('des_eml').AsString);

     Q.Close;
     Q.SQL.Text:='select * from produtos where pro_not='+QuotedStr(QNotas.FieldByName('not_cha').AsString)+' order by pro_des';
     Q.Open;
end;

procedure TFVisualiza.FormShow(Sender: TObject);
begin
     QNotas.Close;
     QNotas.SQL.Text:='select n.*, e.*, d.* from notas n left join destinatarios d on not_dest=des_doc '+
                      'left join emitentes e on not_emit=emi_doc where not_usr='+inttostr(user)+' order by not_imp desc, not_emi desc';
     QNotas.Open;
     QNotas.First;

     MemoE.Clear;
     MemoE.Lines.Add(QNotas.FieldByName('emi_rzo').AsString+' ('+QNotas.FieldByName('emi_fan').AsString+')');
     MemoE.Lines.Add('Documento: '+  QNotas.FieldByName('emi_doc').AsString);
     MemoE.Lines.Add('Endereço: '+  QNotas.FieldByName('emi_end').AsString);
     MemoE.Lines.Add('Cidade: '+  QNotas.FieldByName('emi_cid').AsString);
     MemoE.Lines.Add('Telefone: '+  QNotas.FieldByName('emi_tel').AsString);
     MemoE.Lines.Add('E-mail: '+  QNotas.FieldByName('emi_eml').AsString);

     MemoD.Clear;
     MemoD.Lines.Add(QNotas.FieldByName('des_nom').AsString);
     MemoD.Lines.Add('Documento: '+  QNotas.FieldByName('des_doc').AsString);
     MemoD.Lines.Add('Endereço: '+  QNotas.FieldByName('des_end').AsString);
     MemoD.Lines.Add('Cidade: '+  QNotas.FieldByName('des_cid').AsString);
     MemoD.Lines.Add('Telefone: '+  QNotas.FieldByName('des_tel').AsString);
     MemoD.Lines.Add('E-mail: '+  QNotas.FieldByName('des_eml').AsString);

     Q.Close;
     Q.SQL.Text:='select * from produtos where pro_not='+QuotedStr(QNotas.FieldByName('not_cha').AsString)+' order by pro_des';
     Q.Open;
end;

procedure TFVisualiza.GridNCellClick(Column: TColumn);
begin

end;

procedure TFVisualiza.GridNCellClick(Sender: TObject);
begin

end;


end.

