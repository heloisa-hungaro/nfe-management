unit uremove;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, sqldb, db, FileUtil, Forms, Controls, Graphics, Dialogs,
  StdCtrls, DBGrids, LCLType;

type

  { TFRemove }

  TFRemove = class(TForm)
    BtnCancelar: TButton;
    BtnGerar: TButton;
    D: TDataSource;
    Grid: TDBGrid;
    Q: TSQLQuery;
    QRemove: TSQLQuery;
    Transaction: TSQLTransaction;
    procedure BtnCancelarClick(Sender: TObject);
    procedure BtnGerarClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { private declarations }
  public
    user: integer;
    { public declarations }
  end;

var
  FRemove: TFRemove;

implementation

{$R *.lfm}

{ TFRemove }

procedure TFRemove.BtnCancelarClick(Sender: TObject);
begin
  Q.Close;
  Close;
end;

procedure TFRemove.BtnGerarClick(Sender: TObject);
begin

  if QuestionDlg('Confirmação de Exclusão','Deseja realmente remover a nota do sistema?',
              mtCustom,[mrYes,'Sim',mrNo,'Não'],0) = mrNo then
       Exit;
  QRemove.SQL.Text:='select * from notas where not_cha='+QuotedStr(Q.FieldByName('not_cha').AsString);
  QRemove.Open;
  if (QRemove.RecNo=1) then
  begin
       QRemove.SQL.Text:='delete from produtos where pro_not='+QuotedStr(Q.FieldByName('not_cha').AsString);
       QRemove.ExecSQL;
  end;

  QRemove.SQL.Text:='delete from notas where not_usr='+inttostr(user)+' and not_cha='+QuotedStr(Q.FieldByName('not_cha').AsString);
  QRemove.ExecSQL;
  QRemove.SQL.Text:='select * from notas where not_emit='+IntToStr(Q.FieldByName('not_emit').AsInteger);
  QRemove.Open;
  if (QRemove.RecNo=0) then
  begin
    Q.Close;
    QRemove.SQL.Text:='delete from emitentes where emi_doc='+IntToStr(Q.FieldByName('not_emit').AsInteger);
    QRemove.ExecSQL;
  end;
  QRemove.SQL.Text:='select * from notas where not_dest='+IntToStr(Q.FieldByName('not_dest').AsInteger);
  QRemove.Open;
  if (QRemove.RecNo=0) then
  begin
    Q.Close;
    QRemove.SQL.Text:='delete from destinatarios where des_doc='+IntToStr(Q.FieldByName('not_dest').AsInteger);
    QRemove.ExecSQL;
  end;
  Transaction.Commit;
  Q.Close;
  Q.Open;
end;

procedure TFRemove.FormShow(Sender: TObject);
begin
     Q.Close;
     Q.SQL.Text:='select n.*, e.emi_rzo, d.des_nom from notas n left join destinatarios d on n.not_dest=d.des_doc '+
                      'left join emitentes e on n.not_emit=e.emi_doc where n.not_usr='+inttostr(user)+' order by not_imp desc, not_emi desc';
     Q.Open;
end;

end.

