unit urel;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, sqldb, db, FileUtil, DateTimePicker, Forms, Controls, LCLType, Windows,
  Graphics, Dialogs, StdCtrls, ExtCtrls, DbCtrls, DateUtils, laz2_XMLWrite, laz2_DOM;

type

  { TFRel }

  TFRel = class(TForm)
    BtnGerar: TButton;
    BtnCancelar: TButton;
    ComboDest: TComboBox;
    ComboEmit: TComboBox;
    DataIni: TDateTimePicker;
    DataFim: TDateTimePicker;
    DEmit: TDataSource;
    DDest: TDataSource;
    Label1: TLabel;
    Label2: TLabel;
    LblNome: TLabel;
    LblNome1: TLabel;
    LblNome2: TLabel;
    Periodo: TPanel;
    QEmit: TSQLQuery;
    QDest: TSQLQuery;
    Q: TSQLQuery;
    Diretorio: TSelectDirectoryDialog;
    TipoRel: TRadioGroup;
    procedure BtnCancelarClick(Sender: TObject);
    procedure BtnGerarClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TipoRelClick(Sender: TObject);
  private
    tipo: integer; //1 = geral, 2 = por periodo, 3 - por emitente, 4 - por destinatario
    { private declarations }
  public
    user: integer;
    { public declarations }
  end;

var
  FRel: TFRel;

implementation

{$R *.lfm}

{ TFRel }



procedure TFRel.BtnCancelarClick(Sender: TObject);
begin
  Close;
  QEmit.Close;
  QDest.Close;
end;

procedure TFRel.BtnGerarClick(Sender: TObject);
var caminho,nomearq,aux,vpro,vnot,l1,l2,l3: string; tot_pro,tot_rel: double; qtd_pro, qtd_not: integer;
    xml,xsl,xsd: TXMLDocument; rnode,enode,inode,i2node,i3node,i4node,i5node,tnode,xnode: TDOMNode;
begin
      if (TipoRel.ItemIndex=0) or (TipoRel.ItemIndex=4) then //geral
      begin
           tipo:=1;
           aux:='RELATORIO GERAL';
      end
      else if (TipoRel.ItemIndex=1) or (TipoRel.ItemIndex=5) then //por periodo
      begin
           tipo:=2;
           aux:='RELATORIO POR PERIODO';
      end
      else if (TipoRel.ItemIndex=2) or (TipoRel.ItemIndex=6) then //por emitente
      begin
           tipo:=3;
           aux:='RELATORIO POR EMITENTE';
      end
      else if (TipoRel.ItemIndex=3) or (TipoRel.ItemIndex=7) then //por destinatario
      begin
           tipo:=4;
           aux:='RELATORIO POR DESTINATARIO';
      end;
    Screen.Cursor:=crHourGlass;
    Q.Close;
    Q.SQL.Clear;
    Q.SQL.Text:='select n.*, p.*, d.*, e.*, u.* from notas n left join produtos p on not_cha=pro_not '+
           'left join destinatarios d on not_dest=des_doc left join emitentes e on not_emit=emi_doc '+
           'left join users u on usr_cod=not_usr where not_usr='+inttostr(user)+' ';
    if (TipoRel.ItemIndex=1) or (TipoRel.ItemIndex=5) then //por periodo
       Q.SQL.Text:=Q.SQL.Text+' and not_emi>'+QuotedStr(FormatDateTime('yyyy-mm-dd',DataIni.Date))+
                     ' and not_emi<'+QuotedStr(FormatDateTime('yyyy-mm-dd',DataFim.Date))+' '
    else if (TipoRel.ItemIndex=2) or (TipoRel.ItemIndex=6) then //por emitente
         Q.SQL.Text:=Q.SQL.Text+' and  emi_rzo like '+QuotedStr(ComboEmit.Text)+' '
    else if (TipoRel.ItemIndex=3) or (TipoRel.ItemIndex=7) then //por destinatario
         Q.SQL.Text:=Q.SQL.Text+' and des_nom like '+QuotedStr(ComboDest.Text)+' ';
    Q.SQL.Text:=Q.SQL.Text+' order by not_emi, not_cha';
    Q.Open;
    if (Q.RecNo=0) then
    begin
         Screen.Cursor:=crDefault;
         Application.MessageBox('Não há dados para gerar o relatório desejado','Sem dados',0);
         Exit;
    end;
    nomearq:='REL_D_'+FormatDateTime('dd-mm-yyyy',Now)+'_H_'+FormatDateTime('hh-nn-ss',Now);
    if Diretorio.Execute then
    begin
        try
           caminho := Diretorio.FileName;
           caminho:=caminho+'\'+nomearq;
           tot_pro:=0;
           tot_rel:=0;
           qtd_pro:=0;
           qtd_not:=0;
           //INICIA XML

           xml:=TXMLDocument.Create;
           xml.Encoding:='UTF-8';
           xml.StylesheetHRef:='relatorio.xsl';
           xml.StylesheetType:='text/xsl';
           rNode:=xml.CreateElement('relatorio');
           xml.Appendchild(rNode);
           TDOMElement(rNode).SetAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
           TDOMElement(rNode).SetAttribute('xsi:noNamespaceschemaLocation', 'relatorio.xsd');

           tNode:=xml.CreateTextNode(aux);
           eNode:=xml.CreateElement('tipo');
           rNode.Appendchild(eNode);
           eNode.Appendchild(tNode);

           tNode:=xml.CreateTextNode(DateToStr(Today));
           eNode:=xml.CreateElement('data');
           rNode.Appendchild(eNode);
           eNode.Appendchild(tNode);

           tNode:=xml.CreateTextNode(AnsiUpperCase(Q.FieldByName('usr_log').AsString));
           eNode:=xml.CreateElement('usuario');
           rNode.Appendchild(eNode);
           eNode.Appendchild(tNode);

           if (tipo=2) then
           begin
                eNode:=xml.CreateElement('periodo');
                rNode.Appendchild(eNode);
                iNode:=eNode;

                tNode:=xml.CreateTextNode(DateToStr(DataIni.Date));
                eNode:=xml.CreateElement('inicio');
                iNode.Appendchild(eNode);
                eNode.Appendchild(tNode);

                tNode:=xml.CreateTextNode(DateToStr(DataFim.Date));
                eNode:=xml.CreateElement('fim');
                iNode.Appendchild(eNode);
                eNode.Appendchild(tNode);
           end
           else if (tipo=3) then
           begin
               eNode:=xml.CreateElement('emitente');
               rNode.Appendchild(eNode);
               iNode:=eNode;
               if not (Q.FieldByName('emi_rzo').AsString='') then
               begin
                 tNode:=xml.CreateTextNode(Q.FieldByName('emi_rzo').AsString);
                 eNode:=xml.CreateElement('nome');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);
               end;
               if not (Q.FieldByName('emi_fan').AsString='') then
               begin
                 tNode:=xml.CreateTextNode(Q.FieldByName('emi_fan').AsString);
                 eNode:=xml.CreateElement('fantasia');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);
               end;
               if not (Q.FieldByName('emi_doc').AsString='') then
               begin
                 tNode:=xml.CreateTextNode(Q.FieldByName('emi_doc').AsString);
                 eNode:=xml.CreateElement('documento');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);
               end;
               if not (Q.FieldByName('emi_end').AsString='') then
               begin
                 tNode:=xml.CreateTextNode(Q.FieldByName('emi_end').AsString);
                 eNode:=xml.CreateElement('endereco');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);
               end;
               if not (Q.FieldByName('emi_cid').AsString='') then
               begin
                 tNode:=xml.CreateTextNode(Q.FieldByName('emi_cid').AsString);
                 eNode:=xml.CreateElement('cidade');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);
               end;
               if not (Q.FieldByName('emi_eml').AsString='') then
               begin
                 tNode:=xml.CreateTextNode(Q.FieldByName('emi_eml').AsString);
                 eNode:=xml.CreateElement('email');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);
               end;
               if not (Q.FieldByName('emi_tel').AsString='') then
               begin
                 tNode:=xml.CreateTextNode(Q.FieldByName('emi_tel').AsString);
                 eNode:=xml.CreateElement('telefone');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);
               end;
           end
           else if (tipo=4) then
           begin
               eNode:=xml.CreateElement('destinatario');
               rNode.Appendchild(eNode);
               iNode:=eNode;
               if not (Q.FieldByName('des_nom').AsString='') then
               begin
                 tNode:=xml.CreateTextNode(Q.FieldByName('des_nom').AsString);
                 eNode:=xml.CreateElement('nome');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);
               end;
               if not (Q.FieldByName('des_doc').AsString='') then
               begin
                 tNode:=xml.CreateTextNode(Q.FieldByName('des_doc').AsString);
                 eNode:=xml.CreateElement('documento');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);
               end;
               if not (Q.FieldByName('des_end').AsString='') then
               begin
                 tNode:=xml.CreateTextNode(Q.FieldByName('des_end').AsString);
                 eNode:=xml.CreateElement('endereco');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);
               end;
               if not (Q.FieldByName('des_cid').AsString='') then
               begin
                 tNode:=xml.CreateTextNode(Q.FieldByName('des_cid').AsString);
                 eNode:=xml.CreateElement('cidade');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);
               end;
               if not (Q.FieldByName('des_eml').AsString='') then
               begin
                 tNode:=xml.CreateTextNode(Q.FieldByName('des_eml').AsString);
                 eNode:=xml.CreateElement('email');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);
               end;
               if not (Q.FieldByName('des_tel').AsString='') then
               begin
                 tNode:=xml.CreateTextNode(Q.FieldByName('des_tel').AsString);
                 eNode:=xml.CreateElement('telefone');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);
               end;
           end;
           while not (Q.EOF) do
           begin
                qtd_not:=qtd_not+1;

                eNode:=xml.CreateElement('nota');
                rNode.Appendchild(eNode);
                iNode:=eNode;

                tNode:=xml.CreateTextNode(Q.FieldByName('not_emi').AsString);
                eNode:=xml.CreateElement('emissao');
                iNode.Appendchild(eNode);
                eNode.Appendchild(tNode);
                tNode:=xml.CreateTextNode(Q.FieldByName('not_cha').AsString);
                eNode:=xml.CreateElement('chave');
                iNode.Appendchild(eNode);
                eNode.Appendchild(tNode);
                i2Node:=iNode;
                aux:=Q.FieldByName('not_cha').AsString;
                if (tipo<>3) then
                begin
                   eNode:=xml.CreateElement('emitente');
                   i2Node.Appendchild(eNode);
                   iNode:=eNode;
                   if not (Q.FieldByName('emi_rzo').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('emi_rzo').AsString);
                     eNode:=xml.CreateElement('nome');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('emi_fan').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode('('+Q.FieldByName('emi_fan').AsString+')');
                     eNode:=xml.CreateElement('fantasia');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('emi_doc').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('emi_doc').AsString);
                     eNode:=xml.CreateElement('documento');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('emi_end').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('emi_end').AsString);
                     eNode:=xml.CreateElement('endereco');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('emi_cid').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('emi_cid').AsString);
                     eNode:=xml.CreateElement('cidade');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('emi_eml').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('emi_eml').AsString);
                     eNode:=xml.CreateElement('email');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('emi_tel').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('emi_tel').AsString);
                     eNode:=xml.CreateElement('telefone');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                end;
                if (tipo<>4) then
                begin
                   eNode:=xml.CreateElement('destinatario');
                   i2Node.Appendchild(eNode);
                   iNode:=eNode;
                   if not (Q.FieldByName('des_nom').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('des_nom').AsString);
                     eNode:=xml.CreateElement('nome');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('des_doc').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('des_doc').AsString);
                     eNode:=xml.CreateElement('documento');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('des_end').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('des_end').AsString);
                     eNode:=xml.CreateElement('endereco');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('des_cid').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('des_cid').AsString);
                     eNode:=xml.CreateElement('cidade');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('des_eml').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('des_eml').AsString);
                     eNode:=xml.CreateElement('email');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('des_tel').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('des_tel').AsString);
                     eNode:=xml.CreateElement('telefone');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                end;
               while (aux=Q.FieldByName('not_cha').AsString) and not (Q.EOF) do
                begin
                    qtd_pro:=qtd_pro+1;

                    eNode:=xml.CreateElement('produto');
                    i2Node.Appendchild(eNode);
                    iNode:=eNode;

                    if not (Q.FieldByName('pro_cod').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('pro_cod').AsString);
                     eNode:=xml.CreateElement('codigo');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('pro_des').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('pro_des').AsString);
                     eNode:=xml.CreateElement('descricao');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('pro_uni').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('pro_uni').AsString);
                     eNode:=xml.CreateElement('unidade');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('pro_qtd').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(Q.FieldByName('pro_qtd').AsString);
                     eNode:=xml.CreateElement('quantidade');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('pro_unt').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(FormatFloat('#.00', StrtoFloat(StringReplace(Q.FieldByName('pro_unt').AsString, '.', ',', []))));
                     eNode:=xml.CreateElement('valor_unitario');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;
                   if not (Q.FieldByName('pro_tot').AsString='') then
                   begin
                     tNode:=xml.CreateTextNode(FormatFloat('#.00', StrtoFloat(StringReplace(Q.FieldByName('pro_tot').AsString, '.', ',', []))));
                     eNode:=xml.CreateElement('valor_total');
                     iNode.Appendchild(eNode);
                     eNode.Appendchild(tNode);
                   end;

                    vpro:=Q.FieldByName('not_pro').AsString;
                    vnot:=Q.FieldByName('not_tot').AsString;
                    aux:=Q.FieldByName('not_cha').AsString;
                    Q.Next;
                end;
                tNode:=xml.CreateTextNode(FormatFloat('#.00', StrtoFloat(StringReplace(vpro, '.', ',', []))));
                eNode:=xml.CreateElement('total_produtos');
                i2Node.Appendchild(eNode);
                eNode.Appendchild(tNode);
                tNode:=xml.CreateTextNode(FormatFloat('#.00', StrtoFloat(StringReplace(vnot, '.', ',', []))));
                eNode:=xml.CreateElement('total_nota');
                i2Node.Appendchild(eNode);
                eNode.Appendchild(tNode);
                tot_pro:=tot_pro+StrtoFloat(StringReplace(vpro, '.', ',', []));
                tot_rel:=tot_rel+StrtoFloat(StringReplace(vnot, '.', ',', []));
           end;


           tNode:=xml.CreateTextNode(IntToStr(qtd_pro));
           eNode:=xml.CreateElement('quantidade_produtos');
           rNode.Appendchild(eNode);
           eNode.Appendchild(tNode);
           tNode:=xml.CreateTextNode(FormatFloat('#.00',tot_pro));
           eNode:=xml.CreateElement('total_produtos_relatorio');
           rNode.Appendchild(eNode);
           eNode.Appendchild(tNode);
           tNode:=xml.CreateTextNode(IntToStr(qtd_not));
           eNode:=xml.CreateElement('quantidade_notas');
           rNode.Appendchild(eNode);
           eNode.Appendchild(tNode);
           tNode:=xml.CreateTextNode(FormatFloat('#.00',tot_rel));
           eNode:=xml.CreateElement('total_relatorio');
           rNode.Appendchild(eNode);
           eNode.Appendchild(tNode);

           //FIM XML

           if (TipoRel.ItemIndex>3) then // É HTML
           begin
               //INICIO XSD

               xsd:=TXMLDocument.Create;
               rNode:=xsd.CreateElement('xs:schema');
               xsd.Appendchild(rNode);
               TDOMElement(rNode).SetAttribute('xmlns:xs', 'http://www.w3.org/2001/XMLSchema');
               TDOMElement(rNode).SetAttribute('elementFormDefault', 'qualified');


               iNode:=xsd.CreateElement('xs:element');
               rNode.Appendchild(iNode);
               TDOMElement(iNode).SetAttribute('name', 'relatorio');
               eNode:=xsd.CreateElement('xs:complexType');
               iNode.Appendchild(eNode);
               i2Node:=eNode;

               iNode:=xsd.CreateElement('xs:element');
               i2Node.Appendchild(iNode);
               TDOMElement(iNode).SetAttribute('name', 'tipo');
               TDOMElement(iNode).SetAttribute('type', 'xs:string');

               iNode:=xsd.CreateElement('xs:element');
               i2Node.Appendchild(iNode);
               TDOMElement(iNode).SetAttribute('name', 'data');
               TDOMElement(iNode).SetAttribute('type', 'xs:date');

               iNode:=xsd.CreateElement('xs:element');
               i2Node.Appendchild(iNode);
               TDOMElement(iNode).SetAttribute('name', 'usuario');
               TDOMElement(iNode).SetAttribute('type', 'xs:string');

               if (tipo=2) then
               begin

                    iNode:=xsd.CreateElement('xs:element');
                    i2Node.Appendchild(iNode);
                    TDOMElement(iNode).SetAttribute('name', 'periodo');
                    eNode:=xsd.CreateElement('xs:complexType');
                    iNode.Appendchild(eNode);

                    iNode:=xsd.CreateElement('xs:element');
                    eNode.Appendchild(iNode);
                    TDOMElement(iNode).SetAttribute('name', 'inicio');
                    TDOMElement(iNode).SetAttribute('type', 'xs:date');

                    iNode:=xsd.CreateElement('xs:element');
                    eNode.Appendchild(iNode);
                    TDOMElement(iNode).SetAttribute('name', 'fim');
                    TDOMElement(iNode).SetAttribute('type', 'xs:date');
               end
               else if (tipo=3) then
              begin
                   iNode:=xsd.CreateElement('xs:element');
                   i2Node.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'emitente');
                   eNode:=xsd.CreateElement('xs:complexType');
                   iNode.Appendchild(eNode);

                   iNode:=xsd.CreateElement('xs:element');
                   eNode.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'nome');
                   TDOMElement(iNode).SetAttribute('type', 'xs:string');

                   iNode:=xsd.CreateElement('xs:element');
                   eNode.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'fantasia');
                   TDOMElement(iNode).SetAttribute('type', 'xs:string');

                   iNode:=xsd.CreateElement('xs:element');
                   eNode.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'documento');
                   TDOMElement(iNode).SetAttribute('type', 'xs:integer');

                   iNode:=xsd.CreateElement('xs:element');
                   eNode.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'endereco');
                   TDOMElement(iNode).SetAttribute('type', 'xs:string');

                   iNode:=xsd.CreateElement('xs:element');
                   eNode.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'cidade');
                   TDOMElement(iNode).SetAttribute('type', 'xs:string');

                   iNode:=xsd.CreateElement('xs:element');
                   eNode.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'email');
                   TDOMElement(iNode).SetAttribute('type', 'xs:string');

                   iNode:=xsd.CreateElement('xs:element');
                   eNode.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'telefone');
                   TDOMElement(iNode).SetAttribute('type', 'xs:string');
                end
                else if (tipo=4) then
                begin
                   iNode:=xsd.CreateElement('xs:element');
                   i2Node.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'destinatario');
                   eNode:=xsd.CreateElement('xs:complexType');
                   iNode.Appendchild(eNode);

                   iNode:=xsd.CreateElement('xs:element');
                   eNode.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'nome');
                   TDOMElement(iNode).SetAttribute('type', 'xs:string');

                   iNode:=xsd.CreateElement('xs:element');
                   eNode.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'documento');
                   TDOMElement(iNode).SetAttribute('type', 'xs:integer');

                   iNode:=xsd.CreateElement('xs:element');
                   eNode.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'endereco');
                   TDOMElement(iNode).SetAttribute('type', 'xs:string');

                   iNode:=xsd.CreateElement('xs:element');
                   eNode.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'cidade');
                   TDOMElement(iNode).SetAttribute('type', 'xs:string');

                   iNode:=xsd.CreateElement('xs:element');
                   eNode.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'email');
                   TDOMElement(iNode).SetAttribute('type', 'xs:string');

                   iNode:=xsd.CreateElement('xs:element');
                   eNode.Appendchild(iNode);
                   TDOMElement(iNode).SetAttribute('name', 'telefone');
                   TDOMElement(iNode).SetAttribute('type', 'xs:string');
               end;

               iNode:=xsd.CreateElement('xs:element');
               i2Node.Appendchild(iNode);
               TDOMElement(iNode).SetAttribute('name', 'nota');
               eNode:=xsd.CreateElement('xs:complexType');
               iNode.Appendchild(eNode);
               i3Node:=eNode;

               iNode:=xsd.CreateElement('xs:element');
               i3Node.Appendchild(iNode);
               TDOMElement(iNode).SetAttribute('name', 'emissao');
               TDOMElement(iNode).SetAttribute('type', 'xs:date');

               iNode:=xsd.CreateElement('xs:element');
               i3Node.Appendchild(iNode);
               TDOMElement(iNode).SetAttribute('name', 'chave');
               TDOMElement(iNode).SetAttribute('type', 'xs:string');

               if (tipo<>3) then
               begin
                 iNode:=xsd.CreateElement('xs:element');
                 i3Node.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'emitente');
                 eNode:=xsd.CreateElement('xs:complexType');
                 iNode.Appendchild(eNode);

                 iNode:=xsd.CreateElement('xs:element');
                 eNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'nome');
                 TDOMElement(iNode).SetAttribute('type', 'xs:string');

                 iNode:=xsd.CreateElement('xs:element');
                 eNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'fantasia');
                 TDOMElement(iNode).SetAttribute('type', 'xs:string');

                 iNode:=xsd.CreateElement('xs:element');
                 eNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'documento');
                 TDOMElement(iNode).SetAttribute('type', 'xs:integer');

                 iNode:=xsd.CreateElement('xs:element');
                 eNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'endereco');
                 TDOMElement(iNode).SetAttribute('type', 'xs:string');

                 iNode:=xsd.CreateElement('xs:element');
                 eNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'cidade');
                 TDOMElement(iNode).SetAttribute('type', 'xs:string');

                 iNode:=xsd.CreateElement('xs:element');
                 eNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'email');
                 TDOMElement(iNode).SetAttribute('type', 'xs:string');

                 iNode:=xsd.CreateElement('xs:element');
                 eNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'telefone');
                 TDOMElement(iNode).SetAttribute('type', 'xs:string');
               end;
               if (tipo<>4) then
               begin
                  iNode:=xsd.CreateElement('xs:element');
                  i3Node.Appendchild(iNode);
                  TDOMElement(iNode).SetAttribute('name', 'destinatario');
                  eNode:=xsd.CreateElement('xs:complexType');
                  iNode.Appendchild(eNode);

                  iNode:=xsd.CreateElement('xs:element');
                  eNode.Appendchild(iNode);
                  TDOMElement(iNode).SetAttribute('name', 'nome');
                  TDOMElement(iNode).SetAttribute('type', 'xs:string');

                  iNode:=xsd.CreateElement('xs:element');
                  eNode.Appendchild(iNode);
                  TDOMElement(iNode).SetAttribute('name', 'documento');
                  TDOMElement(iNode).SetAttribute('type', 'xs:integer');

                  iNode:=xsd.CreateElement('xs:element');
                  eNode.Appendchild(iNode);
                  TDOMElement(iNode).SetAttribute('name', 'endereco');
                  TDOMElement(iNode).SetAttribute('type', 'xs:string');

                  iNode:=xsd.CreateElement('xs:element');
                  eNode.Appendchild(iNode);
                  TDOMElement(iNode).SetAttribute('name', 'cidade');
                  TDOMElement(iNode).SetAttribute('type', 'xs:string');

                  iNode:=xsd.CreateElement('xs:element');
                  eNode.Appendchild(iNode);
                  TDOMElement(iNode).SetAttribute('name', 'email');
                  TDOMElement(iNode).SetAttribute('type', 'xs:string');

                  iNode:=xsd.CreateElement('xs:element');
                  eNode.Appendchild(iNode);
                  TDOMElement(iNode).SetAttribute('name', 'telefone');
                  TDOMElement(iNode).SetAttribute('type', 'xs:string');
               end;

                 iNode:=xsd.CreateElement('xs:element');
                 i3Node.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'produto');
                 eNode:=xsd.CreateElement('xs:complexType');
                 iNode.Appendchild(eNode);

                 iNode:=xsd.CreateElement('xs:element');
                 eNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'codigo');
                 TDOMElement(iNode).SetAttribute('type', 'xs:string');

                 iNode:=xsd.CreateElement('xs:element');
                 eNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'descricao');
                 TDOMElement(iNode).SetAttribute('type', 'xs:integer');

                 iNode:=xsd.CreateElement('xs:element');
                 eNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'unidade');
                 TDOMElement(iNode).SetAttribute('type', 'xs:integer');

                 iNode:=xsd.CreateElement('xs:element');
                 eNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'quantidade');
                 TDOMElement(iNode).SetAttribute('type', 'xs:decimal');

                 iNode:=xsd.CreateElement('xs:element');
                 eNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'valor_unitario');
                 TDOMElement(iNode).SetAttribute('type', 'xs:decimal');

                 iNode:=xsd.CreateElement('xs:element');
                 eNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'valor_total');
                 TDOMElement(iNode).SetAttribute('type', 'xs:decimal');

                 iNode:=xsd.CreateElement('xs:element');
                 eNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'telefone');
                 TDOMElement(iNode).SetAttribute('type', 'xs:string');


                 iNode:=xsd.CreateElement('xs:element');
                 i3Node.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'total_produtos');
                 TDOMElement(iNode).SetAttribute('type', 'xs:decimal');

                 iNode:=xsd.CreateElement('xs:element');
                 i3Node.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('name', 'total_nota');
                 TDOMElement(iNode).SetAttribute('type', 'xs:decimal');

               iNode:=xsd.CreateElement('xs:element');
               i2Node.Appendchild(iNode);
               TDOMElement(iNode).SetAttribute('name', 'quantidade_notas');
               TDOMElement(iNode).SetAttribute('type', 'xs:integer');

               iNode:=xsd.CreateElement('xs:element');
               i2Node.Appendchild(iNode);
               TDOMElement(iNode).SetAttribute('name', 'total_produtos_relatorio');
               TDOMElement(iNode).SetAttribute('type', 'xs:decimal');

               iNode:=xsd.CreateElement('xs:element');
               i2Node.Appendchild(iNode);
               TDOMElement(iNode).SetAttribute('name', 'quantidade_produtos');
               TDOMElement(iNode).SetAttribute('type', 'xs:integer');

               iNode:=xsd.CreateElement('xs:element');
               i2Node.Appendchild(iNode);
               TDOMElement(iNode).SetAttribute('name', 'total_relatorio');
               TDOMElement(iNode).SetAttribute('type', 'xs:decimal');

               //FIM XSD


                //INICIO XSL

                 xsl:=TXMLDocument.Create;
                 rNode:=xsl.CreateElement('xsl:stylesheet');
                 xsl.Appendchild(rNode);
                 TDOMElement(rNode).SetAttribute('version','1.0');
                 TDOMElement(rNode).SetAttribute('xmlns:xsl', 'http://www.w3.org/1999/XSL/Transform');


                 iNode:=xsl.CreateElement('xsl:output');
                 rNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('method', 'html');

                 iNode:=xsl.CreateElement('xsl:template');
                 rNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('match', 'relatorio');

                 i2Node:=xsl.CreateElement('html');
                 iNode.Appendchild(i2Node);

                 i3Node:=xsl.CreateElement('head');
                 i2Node.Appendchild(i3Node);

                 l1:='body { border: 1px solid black; width: 70%; margin: 10px auto; font-family: Arial,sans-serif; font-size: 10pt } h1 { font-size: 14pt; text-align: center } h2 { font-size: 12pt } h3 { font-size: 8pt; text-align: center } h1,h2,h3,p { margin: 10px } ';
                 l2:='.titulo { width: 96%; margin: 10px auto; border-style: dotted } .menor { display: none; } .bloco { display: inline-block; border-right: 1px solid black; text-align: center; margin: 0 2px } ';
                 l3:='@media screen and (max-width: 700px) { body { width: 400px; } } @media screen and (max-width: 1200px) { .menor { display: block; } .bloco { display: none; } }';
                 tNode:=xsl.CreateTextNode(l1+l2+l3);
                 eNode:=xsl.CreateElement('style');
                 i3Node.Appendchild(eNode);
                 eNode.Appendchild(tNode);

                 i3Node:=xsl.CreateElement('body');
                 i2Node.Appendchild(i3Node);

                 eNode:=xsl.CreateElement('h1');
                 i3Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'tipo');

                 i4Node:=xsl.CreateElement('h3');
                 i3Node.Appendchild(i4Node);


                 tNode:=xsl.CreateTextNode('Emitido em ');
                 i4Node.Appendchild(tNode);

                 eNode:=xsl.CreateElement('xsl:value-of');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('select', 'data');

                 tNode:=xsl.CreateTextNode(' por ');
                 i4Node.Appendchild(tNode);

                 eNode:=xsl.CreateElement('xsl:value-of');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('select', 'usuario');

                 if (tipo=2) then
                 begin
                      i4Node:=xsl.CreateElement('h2');
                      i3Node.Appendchild(i4Node);
                      TDOMElement(i4Node).SetAttribute('style', 'color: #990000');

                      eNode:=xsl.CreateElement('xsl:apply-templates');
                      i4Node.Appendchild(eNode);
                      TDOMElement(eNode).SetAttribute('select', 'periodo');

                 end;

                 eNode:=xsl.CreateElement('p');
                 i3Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('style', 'border-bottom: 2px solid black');

                 if (tipo=3) then
                 begin
                     tNode:=xsl.CreateTextNode('Emitente:');
                     eNode:=xsl.CreateElement('h2');
                     i3Node.Appendchild(eNode);
                     eNode.Appendchild(tNode);

                     eNode:=xsl.CreateElement('div');
                     i3Node.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('class', 'titulo');
                     tNode:=xsl.CreateElement('xsl:apply-templates');
                     eNode.Appendchild(tNode);
                     TDOMElement(tNode).SetAttribute('select', 'emitente');

					 eNode:=xsl.CreateElement('p');
                     i3Node.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('style', 'margin: 30px 10px; border-bottom: 5px double #990000');
                     eNode:=xsl.CreateElement('p');
                     i3Node.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('style', 'margin: 30px 10px; border-bottom: 5px double #990000');
                 end
                 else if (tipo=4) then
                 begin

                     tNode:=xsl.CreateTextNode('Destinatario:');
                     eNode:=xsl.CreateElement('h2');
                     i3Node.Appendchild(eNode);
                     eNode.Appendchild(tNode);

                     eNode:=xsl.CreateElement('div');
                     i3Node.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('class', 'titulo');
                     tNode:=xsl.CreateElement('xsl:apply-templates');
                     eNode.Appendchild(tNode);
                     TDOMElement(tNode).SetAttribute('select', 'destinatario');

		     eNode:=xsl.CreateElement('p');
                     i3Node.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('style', 'margin: 30px 10px; border-bottom: 5px double #990000');
                     eNode:=xsl.CreateElement('p');
                     i3Node.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('style', 'margin: 30px 10px; border-bottom: 5px double #990000');
                 end;

                 i4Node:=xsl.CreateElement('xsl:for-each');
                 i3Node.Appendchild(i4Node);
                 TDOMElement(i4Node).SetAttribute('select', 'nota');

                 i5Node:=xsl.CreateElement('div');
                 i4Node.Appendchild(i5Node);
                 TDOMElement(i5Node).SetAttribute('class', 'titulo');

                 eNode:=xsl.CreateElement('p');
                 i5Node.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Nota emitida em ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'emissao');

                 eNode:=xsl.CreateElement('p');
                 i5Node.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Chave de acesso: ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'chave');

                 eNode:=xsl.CreateElement('p');
                 i5Node.Appendchild(eNode);
                 if (tipo<>3) then
                 begin
                     tNode:=xsl.CreateTextNode('Emitente:');
                     eNode:=xsl.CreateElement('h2');
                     i4Node.Appendchild(eNode);
                     eNode.Appendchild(tNode);

                     eNode:=xsl.CreateElement('div');
                     i4Node.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('class', 'titulo');
                     tNode:=xsl.CreateElement('xsl:apply-templates');
                     eNode.Appendchild(tNode);
                     TDOMElement(tNode).SetAttribute('select', 'emitente');
                 end;
                 if (tipo<>4) then
                 begin

                     tNode:=xsl.CreateTextNode('Destinatario:');
                     eNode:=xsl.CreateElement('h2');
                     i4Node.Appendchild(eNode);
                     eNode.Appendchild(tNode);

                     eNode:=xsl.CreateElement('div');
                     i4Node.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('class', 'titulo');
                     tNode:=xsl.CreateElement('xsl:apply-templates');
                     eNode.Appendchild(tNode);
                     TDOMElement(tNode).SetAttribute('select', 'destinatario');
                 end;

                 tNode:=xsl.CreateTextNode('Produtos');
                 eNode:=xsl.CreateElement('h2');
                 i4Node.Appendchild(eNode);
                 eNode.Appendchild(tNode);

                 i5Node:=xsl.CreateElement('div');
                 i4Node.Appendchild(i5Node);
                 TDOMElement(i5Node).SetAttribute('class', 'titulo');
                 TDOMElement(i5Node).SetAttribute('style', 'padding: 4px 0');

                 iNode:=xsl.CreateElement('div');
                 i5Node.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('class', 'bloco');
                 TDOMElement(iNode).SetAttribute('style', 'width: 12%; padding-left: 4px');

                 tNode:=xsl.CreateTextNode('COD');
                 eNode:=xsl.CreateElement('b');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);

                 iNode:=xsl.CreateElement('div');
                 i5Node.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('class', 'bloco');
                 TDOMElement(iNode).SetAttribute('style', 'width: 50%');

                 tNode:=xsl.CreateTextNode('DESC');
                 eNode:=xsl.CreateElement('b');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);

                 iNode:=xsl.CreateElement('div');
                 i5Node.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('class', 'bloco');
                 TDOMElement(iNode).SetAttribute('style', 'width: 7%');

                 tNode:=xsl.CreateTextNode('UNI');
                 eNode:=xsl.CreateElement('b');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);

                 iNode:=xsl.CreateElement('div');
                 i5Node.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('class', 'bloco');
                 TDOMElement(iNode).SetAttribute('style', 'width: 7%');

                 tNode:=xsl.CreateTextNode('QTDE');
                 eNode:=xsl.CreateElement('b');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);

                 iNode:=xsl.CreateElement('div');
                 i5Node.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('class', 'bloco');
                 TDOMElement(iNode).SetAttribute('style', 'width: 9%; padding-right: 4px');

                 tNode:=xsl.CreateTextNode('R$ UNIT');
                 eNode:=xsl.CreateElement('b');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);

                 iNode:=xsl.CreateElement('div');
                 i5Node.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('class', 'bloco');
                 TDOMElement(iNode).SetAttribute('style', 'width: 9%; border-right: 0px');

                 tNode:=xsl.CreateTextNode('R$ TOT');
                 eNode:=xsl.CreateElement('b');
                 iNode.Appendchild(eNode);
                 eNode.Appendchild(tNode);

                 iNode:=xsl.CreateElement('xsl:for-each');
                 i5Node.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('select', 'produto');

                 i2Node:=xsl.CreateElement('div');
                 iNode.Appendchild(i2Node);
                 TDOMElement(i2Node).SetAttribute('class', 'menor');

                 eNode:=xsl.CreateElement('p');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('b');
                 eNode.Appendchild(tNode);
                 xNode:=xsl.CreateTextNode('Codigo');
                 tNode.Appendchild(xNode);
                 tNode:=xsl.CreateTextNode(': ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'codigo');

                 eNode:=xsl.CreateElement('p');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('b');
                 eNode.Appendchild(tNode);
                 xNode:=xsl.CreateTextNode('Descricao');
                 tNode.Appendchild(xNode);
                 tNode:=xsl.CreateTextNode(': ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'descricao');

                 eNode:=xsl.CreateElement('p');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('b');
                 eNode.Appendchild(tNode);
                 xNode:=xsl.CreateTextNode('Unidade');
                 tNode.Appendchild(xNode);
                 tNode:=xsl.CreateTextNode(': ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'unidade');

                 eNode:=xsl.CreateElement('p');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('b');
                 eNode.Appendchild(tNode);
                 xNode:=xsl.CreateTextNode('Quantidade');
                 tNode.Appendchild(xNode);
                 tNode:=xsl.CreateTextNode(': ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'quantidade');

                 eNode:=xsl.CreateElement('p');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('b');
                 eNode.Appendchild(tNode);
                 xNode:=xsl.CreateTextNode('Valor unitario');
                 tNode.Appendchild(xNode);
                 tNode:=xsl.CreateTextNode(': ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'valor_unitario');

                 eNode:=xsl.CreateElement('p');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('b');
                 eNode.Appendchild(tNode);
                 xNode:=xsl.CreateTextNode('Valor total');
                 tNode.Appendchild(xNode);
                 tNode:=xsl.CreateTextNode(': ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'valor_total');

                 eNode:=xsl.CreateElement('p');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('style', 'border-bottom: 1px solid black');

                 i2Node:=xsl.CreateElement('div');
                 iNode.Appendchild(i2Node);
                 TDOMElement(i2Node).SetAttribute('class', 'bloco');
                 TDOMElement(i2Node).SetAttribute('style', 'width: 12%; text-align: left; padding-left: 4px');
                 tNode:=xsl.CreateTextNode('   ');
                 i2Node.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'codigo');

                 i2Node:=xsl.CreateElement('div');
                 iNode.Appendchild(i2Node);
                 TDOMElement(i2Node).SetAttribute('class', 'bloco');
                 TDOMElement(i2Node).SetAttribute('style', 'width: 50%; text-align: left');
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'descricao');

                 i2Node:=xsl.CreateElement('div');
                 iNode.Appendchild(i2Node);
                 TDOMElement(i2Node).SetAttribute('class', 'bloco');
                 TDOMElement(i2Node).SetAttribute('style', 'width: 7%; text-align: center');
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'unidade');

                 i2Node:=xsl.CreateElement('div');
                 iNode.Appendchild(i2Node);
                 TDOMElement(i2Node).SetAttribute('class', 'bloco');
                 TDOMElement(i2Node).SetAttribute('style', 'width: 7%; text-align: center');
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'quantidade');

                 i2Node:=xsl.CreateElement('div');
                 iNode.Appendchild(i2Node);
                 TDOMElement(i2Node).SetAttribute('class', 'bloco');
                 TDOMElement(i2Node).SetAttribute('style', 'width: 9%; text-align: right; padding-right: 4px');
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'valor_unitario');

                 i2Node:=xsl.CreateElement('div');
                 iNode.Appendchild(i2Node);
                 TDOMElement(i2Node).SetAttribute('class', 'bloco');
                 TDOMElement(i2Node).SetAttribute('style', 'width: 9%; text-align: right; border-right: 0px');
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'valor_total');

                 iNode:=xsl.CreateElement('div');
                 i4Node.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('class', 'titulo');

                 eNode:=xsl.CreateElement('p');
                 iNode.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Valor total dos produtos: ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'total_produtos');

                 eNode:=xsl.CreateElement('p');
                 iNode.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Valor total da nota: ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'total_nota');

                 eNode:=xsl.CreateElement('p');
                 iNode.Appendchild(eNode);

                 eNode:=xsl.CreateElement('p');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('style', 'margin: 30px 10px; border-bottom: 5px double #990000');

                 eNode:=xsl.CreateElement('p');
                 i3Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('style', 'margin: 30px 10px; border-bottom: 5px double #990000');

                 eNode:=xsl.CreateElement('p');
                 i3Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('style', 'margin: 30px 10px; border-bottom: 5px double #990000');

                 iNode:=xsl.CreateElement('div');
                 i3Node.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('class', 'titulo');

                 eNode:=xsl.CreateElement('p');
                 iNode.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Quantidade de produtos: ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'quantidade_produtos');

                 eNode:=xsl.CreateElement('p');
                 iNode.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Valor total dos produtos: ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'total_produtos_relatorio');

                 eNode:=xsl.CreateElement('p');
                 iNode.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Quantidade de notas: ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'quantidade_notas');

                 eNode:=xsl.CreateElement('p');
                 iNode.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Valor total das notas: ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'total_relatorio');

                 eNode:=xsl.CreateElement('p');
                 iNode.Appendchild(eNode);


                 iNode:=xsl.CreateElement('xsl:template');
                 rNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('match', 'emitente');

                 i2Node:=xsl.CreateElement('p');
                 iNode.Appendchild(i2Node);
                 eNode:=xsl.CreateElement('b');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Nome / Razão Social');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode(': ');
                 i2Node.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'nome');
                 tNode:=xsl.CreateTextNode(' ');
                 i2Node.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'fantasia');

                 i2Node:=xsl.CreateElement('p');
                 iNode.Appendchild(i2Node);
                 eNode:=xsl.CreateElement('b');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('CPF / CNPJ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode(': ');
                 i2Node.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'documento');

                 i2Node:=xsl.CreateElement('p');
                 iNode.Appendchild(i2Node);
                 eNode:=xsl.CreateElement('b');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Endereco');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode(': ');
                 i2Node.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'endereco');

                 i2Node:=xsl.CreateElement('p');
                 iNode.Appendchild(i2Node);
                 eNode:=xsl.CreateElement('b');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Cidade');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode(': ');
                 i2Node.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'cidade');

                 i2Node:=xsl.CreateElement('p');
                 iNode.Appendchild(i2Node);
                 eNode:=xsl.CreateElement('b');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('E-mail');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode(': ');
                 i2Node.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'email');

                 i2Node:=xsl.CreateElement('p');
                 iNode.Appendchild(i2Node);
                 eNode:=xsl.CreateElement('b');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Telefone');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode(': ');
                 i2Node.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'telefone');


                 iNode:=xsl.CreateElement('xsl:template');
                 rNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('match', 'destinatario');

                 i2Node:=xsl.CreateElement('p');
                 iNode.Appendchild(i2Node);
                 eNode:=xsl.CreateElement('b');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Nome / Razão Social');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode(': ');
                 i2Node.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'nome');

                 i2Node:=xsl.CreateElement('p');
                 iNode.Appendchild(i2Node);
                 eNode:=xsl.CreateElement('b');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('CPF / CNPJ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode(': ');
                 i2Node.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'documento');

                 i2Node:=xsl.CreateElement('p');
                 iNode.Appendchild(i2Node);
                 eNode:=xsl.CreateElement('b');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Endereco');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode(': ');
                 i2Node.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'endereco');

                 i2Node:=xsl.CreateElement('p');
                 iNode.Appendchild(i2Node);
                 eNode:=xsl.CreateElement('b');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Cidade');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode(': ');
                 i2Node.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'cidade');

                 i2Node:=xsl.CreateElement('p');
                 iNode.Appendchild(i2Node);
                 eNode:=xsl.CreateElement('b');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('E-mail');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode(': ');
                 i2Node.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'email');

                 i2Node:=xsl.CreateElement('p');
                 iNode.Appendchild(i2Node);
                 eNode:=xsl.CreateElement('b');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateTextNode('Telefone');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode(': ');
                 i2Node.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'telefone');

                 if (tipo=2) then
                 begin
                     iNode:=xsl.CreateElement('xsl:template');
                     rNode.Appendchild(iNode);
                     TDOMElement(iNode).SetAttribute('match', 'periodo');

                     tNode:=xsl.CreateTextNode('Notas emitidas entre ');
                     iNode.Appendchild(tNode);

                     eNode:=xsl.CreateElement('xsl:value-of');
                     iNode.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('select', 'inicio');

                     tNode:=xsl.CreateTextNode(' e ');
                     iNode.Appendchild(tNode);

                     eNode:=xsl.CreateElement('xsl:value-of');
                     iNode.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('select', 'fim');
                 end;


           end
           else // É PDF
           begin
                 xsl:=TXMLDocument.Create;
                 xsl.Encoding:='UTF-8';
                 rNode:=xsl.CreateElement('xsl:stylesheet');
                 xsl.Appendchild(rNode);
                 TDOMElement(rNode).SetAttribute('version','1.0');
                 TDOMElement(rNode).SetAttribute('xmlns:xsl', 'http://www.w3.org/1999/XSL/Transform');
                 TDOMElement(rNode).SetAttribute('xmlns:fo', 'http://www.w3.org/1999/XSL/Format');

                 iNode:=xsl.CreateElement('xsl:template');
                 rNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('match', '/');

                 i2Node:=xsl.CreateElement('fo:root');
                 iNode.Appendchild(i2Node);
                 TDOMElement(i2Node).SetAttribute('xmlns:fo', 'http://www.w3.org/1999/XSL/Format');

                 i3Node:=xsl.CreateElement('fo:layout-master-set');
                 i2Node.Appendchild(i3Node);

                 i4Node:=xsl.CreateElement('fo:simple-page-master');
                 i3Node.Appendchild(i4Node);
                 TDOMElement(i4Node).SetAttribute('master-name', 'relpdf');
                 TDOMElement(i4Node).SetAttribute('page-height', '29.7cm');
                 TDOMElement(i4Node).SetAttribute('page-width', '21cm');
                 TDOMElement(i4Node).SetAttribute('margin-top', '2cm');
                 TDOMElement(i4Node).SetAttribute('margin-bottom', '2cm');
                 TDOMElement(i4Node).SetAttribute('margin-left', '2.5cm');
                 TDOMElement(i4Node).SetAttribute('margin-right', '2.5cm');

                 eNode:=xsl.CreateElement('fo:region-body');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('margin-top', '0.5cm');

                 eNode:=xsl.CreateElement('fo:region-before');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('extent', '2cm');

                 i3Node:=xsl.CreateElement('fo:page-sequence');
                 i2Node.Appendchild(i3Node);
                 TDOMElement(i3Node).SetAttribute('master-reference', 'relpdf');
                 TDOMElement(i3Node).SetAttribute('initial-page-number', '1');

                 i4Node:=xsl.CreateElement('fo:static-content');
                 i3Node.Appendchild(i4Node);
                 TDOMElement(i4Node).SetAttribute('flow-name', 'xsl-region-before');

                 i5Node:=xsl.CreateElement('fo:block');
                 i4Node.Appendchild(i5Node);
                 TDOMElement(i5Node).SetAttribute('text-align', 'end');
                 TDOMElement(i5Node).SetAttribute('font-size', '12pt');
                 TDOMElement(i5Node).SetAttribute('font-family', 'sans-serif');

                 eNode:=xsl.CreateElement('fo:page-number');
                 i5Node.Appendchild(eNode);

                 i4Node:=xsl.CreateElement('fo:flow');
                 i3Node.Appendchild(i4Node);
                 TDOMElement(i4Node).SetAttribute('flow-name', 'xsl-region-body');

                 eNode:=xsl.CreateElement('xsl:apply-templates');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('select', 'relatorio');

                 iNode:=xsl.CreateElement('xsl:template');
                 rNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('match', 'relatorio');

                 i2Node:=xsl.CreateElement('fo:block');
                 iNode.Appendchild(i2Node);
                 TDOMElement(i2Node).SetAttribute('font-size', '16pt');
                 TDOMElement(i2Node).SetAttribute('font-family', 'sans-serif');
                 TDOMElement(i2Node).SetAttribute('font-weight', '400');
                 TDOMElement(i2Node).SetAttribute('line-height', '24pt');
                 TDOMElement(i2Node).SetAttribute('space-after.optimum', '5pt');
                 TDOMElement(i2Node).SetAttribute('text-align', 'center');
                 TDOMElement(i2Node).SetAttribute('padding-top', '2pt');

                 eNode:=xsl.CreateElement('xsl:value-of');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('select', 'tipo');

                 i2Node:=xsl.CreateElement('fo:block');
                 iNode.Appendchild(i2Node);
                 TDOMElement(i2Node).SetAttribute('font-size', '8pt');
                 TDOMElement(i2Node).SetAttribute('font-family', 'sans-serif');
                 TDOMElement(i2Node).SetAttribute('space-after.optimum', '5pt');
                 TDOMElement(i2Node).SetAttribute('text-align', 'center');
                 TDOMElement(i2Node).SetAttribute('padding-bottom', '8pt');

                 i3Node:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(i3Node);

                 tNode:=xsl.CreateTextNode('Emitido em ');
                 i3Node.Appendchild(tNode);
                 eNode:=xsl.CreateElement('xsl:value-of');
                 i3Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('select', 'data');
                 tNode:=xsl.CreateTextNode(' por ');
                 i3Node.Appendchild(tNode);
                 eNode:=xsl.CreateElement('xsl:value-of');
                 i3Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('select', 'usuario');

                 if (tipo=2) then
                 begin
                     eNode:=xsl.CreateElement('xsl:apply-templates');
                     iNode.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('select', 'periodo');
                 end;

                 i2Node:=xsl.CreateElement('fo:block');
                 iNode.Appendchild(i2Node);
                 TDOMElement(i2Node).SetAttribute('font-size', '0pt');
                 TDOMElement(i2Node).SetAttribute('border-bottom-width', '1pt');
                 TDOMElement(i2Node).SetAttribute('border-bottom-style', 'solid');
                 TDOMElement(i2Node).SetAttribute('border-bottom-color', 'black');

                 if (tipo=3) then
                 begin
                     i3Node:=xsl.CreateElement('fo:block');
                     iNode.Appendchild(i3Node);
                     TDOMElement(i3Node).SetAttribute('font-size', '12pt');
                     TDOMElement(i3Node).SetAttribute('font-family', 'sans-serif');
                     TDOMElement(i3Node).SetAttribute('font-weight', 'bold');
                     TDOMElement(i3Node).SetAttribute('color', 'black');
                     TDOMElement(i3Node).SetAttribute('space-after.optimum', '5pt');
                     TDOMElement(i3Node).SetAttribute('text-align', 'left');
                     TDOMElement(i3Node).SetAttribute('margin-top', '10pt');

                     tNode:=xsl.CreateTextNode('Emitente:');
                     i3Node.Appendchild(tNode);

                     eNode:=xsl.CreateElement('xsl:apply-templates');
                     iNode.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('select', 'emitente');

                     eNode:=xsl.CreateElement('fo:block');
                     iNode.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('font-size', '0pt');
                     TDOMElement(eNode).SetAttribute('margin-top', '15pt');
                     TDOMElement(eNode).SetAttribute('border-bottom-width', '3pt');
                     TDOMElement(eNode).SetAttribute('border-bottom-style', 'double');
                     TDOMElement(eNode).SetAttribute('border-bottom-color', '#990000');

                     eNode:=xsl.CreateElement('fo:block');
                     iNode.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('font-size', '0pt');
                     TDOMElement(eNode).SetAttribute('margin-top', '15pt');
                     TDOMElement(eNode).SetAttribute('border-bottom-width', '3pt');
                     TDOMElement(eNode).SetAttribute('border-bottom-style', 'double');
                     TDOMElement(eNode).SetAttribute('border-bottom-color', '#990000');
                     end
                 else if (tipo=4) then
                 begin

                     i3Node:=xsl.CreateElement('fo:block');
                     iNode.Appendchild(i3Node);
                     TDOMElement(i3Node).SetAttribute('font-size', '12pt');
                     TDOMElement(i3Node).SetAttribute('font-family', 'sans-serif');
                     TDOMElement(i3Node).SetAttribute('font-weight', 'bold');
                     TDOMElement(i3Node).SetAttribute('color', 'black');
                     TDOMElement(i3Node).SetAttribute('space-after.optimum', '5pt');
                     TDOMElement(i3Node).SetAttribute('text-align', 'left');
                     TDOMElement(i3Node).SetAttribute('margin-top', '10pt');

                     tNode:=xsl.CreateTextNode('Destinatario:');
                     i3Node.Appendchild(tNode);

                     eNode:=xsl.CreateElement('xsl:apply-templates');
                     iNode.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('select', 'destinatario');

                     eNode:=xsl.CreateElement('fo:block');
                     iNode.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('font-size', '0pt');
                     TDOMElement(eNode).SetAttribute('margin-top', '15pt');
                     TDOMElement(eNode).SetAttribute('border-bottom-width', '3pt');
                     TDOMElement(eNode).SetAttribute('border-bottom-style', 'double');
                     TDOMElement(eNode).SetAttribute('border-bottom-color', '#990000');

                     eNode:=xsl.CreateElement('fo:block');
                     iNode.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('font-size', '0pt');
                     TDOMElement(eNode).SetAttribute('margin-top', '15pt');
                     TDOMElement(eNode).SetAttribute('border-bottom-width', '3pt');
                     TDOMElement(eNode).SetAttribute('border-bottom-style', 'double');
                     TDOMElement(eNode).SetAttribute('border-bottom-color', '#990000');
                 end;

                 i2Node:=xsl.CreateElement('xsl:for-each');
                 iNode.Appendchild(i2Node);
                 TDOMElement(i2Node).SetAttribute('select', 'nota');

                 i3Node:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(i3Node);
                 TDOMElement(i3Node).SetAttribute('font-size', '10pt');
                 TDOMElement(i3Node).SetAttribute('font-family', 'sans-serif');
                 TDOMElement(i3Node).SetAttribute('color', '#303030');
                 TDOMElement(i3Node).SetAttribute('space-after.optimum', '5pt');
                 TDOMElement(i3Node).SetAttribute('text-align', 'left');
                 TDOMElement(i3Node).SetAttribute('padding', '4pt');
                 TDOMElement(i3Node).SetAttribute('margin-top', '15pt');
                 TDOMElement(i3Node).SetAttribute('border-width', '3pt');
                 TDOMElement(i3Node).SetAttribute('border-style', 'dotted');
                 TDOMElement(i3Node).SetAttribute('border-color', 'black');

                 i4Node:=xsl.CreateElement('fo:inline');
                 i3Node.Appendchild(i4Node);

                 tNode:=xsl.CreateTextNode('Nota emitida em ');
                 i4Node.Appendchild(tNode);
                 eNode:=xsl.CreateElement('xsl:value-of');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('select', 'emissao');
                 tNode:=xsl.CreateElement('fo:block');
                 i4Node.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode('Chave de acesso: ');
                 i4Node.Appendchild(tNode);
                 eNode:=xsl.CreateElement('xsl:value-of');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('select', 'chave');

                 if (tipo<>3) then
                 begin
                     i3Node:=xsl.CreateElement('fo:block');
                     i2Node.Appendchild(i3Node);
                     TDOMElement(i3Node).SetAttribute('font-size', '12pt');
                     TDOMElement(i3Node).SetAttribute('font-family', 'sans-serif');
                     TDOMElement(i3Node).SetAttribute('font-weight', 'bold');
                     TDOMElement(i3Node).SetAttribute('color', 'black');
                     TDOMElement(i3Node).SetAttribute('space-after.optimum', '5pt');
                     TDOMElement(i3Node).SetAttribute('text-align', 'left');
                     TDOMElement(i3Node).SetAttribute('margin-top', '10pt');

                     tNode:=xsl.CreateTextNode('Emitente:');
                     i3Node.Appendchild(tNode);

                     eNode:=xsl.CreateElement('xsl:apply-templates');
                     i2Node.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('select', 'emitente');
                 end;
                 if (tipo<>4) then
                 begin
                     i3Node:=xsl.CreateElement('fo:block');
                     i2Node.Appendchild(i3Node);
                     TDOMElement(i3Node).SetAttribute('font-size', '12pt');
                     TDOMElement(i3Node).SetAttribute('font-family', 'sans-serif');
                     TDOMElement(i3Node).SetAttribute('font-weight', 'bold');
                     TDOMElement(i3Node).SetAttribute('color', 'black');
                     TDOMElement(i3Node).SetAttribute('space-after.optimum', '5pt');
                     TDOMElement(i3Node).SetAttribute('text-align', 'left');
                     TDOMElement(i3Node).SetAttribute('margin-top', '10pt');

                     tNode:=xsl.CreateTextNode('Destinatario:');
                     i3Node.Appendchild(tNode);

                     eNode:=xsl.CreateElement('xsl:apply-templates');
                     i2Node.Appendchild(eNode);
                     TDOMElement(eNode).SetAttribute('select', 'destinatario');
                 end;

                 i3Node:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(i3Node);
                 TDOMElement(i3Node).SetAttribute('font-size', '12pt');
                 TDOMElement(i3Node).SetAttribute('font-family', 'sans-serif');
                 TDOMElement(i3Node).SetAttribute('font-weight', 'bold');
                 TDOMElement(i3Node).SetAttribute('color', 'black');
                 TDOMElement(i3Node).SetAttribute('space-after.optimum', '5pt');
                 TDOMElement(i3Node).SetAttribute('text-align', 'left');
                 TDOMElement(i3Node).SetAttribute('margin-top', '10pt');

                 tNode:=xsl.CreateTextNode('Produtos:');
                 i3Node.Appendchild(tNode);

                 i3Node:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(i3Node);
                 TDOMElement(i3Node).SetAttribute('font-size', '7pt');
                 TDOMElement(i3Node).SetAttribute('font-family', 'sans-serif');
                 TDOMElement(i3Node).SetAttribute('color', '#303030');
                 TDOMElement(i3Node).SetAttribute('border-width', '3pt');
                 TDOMElement(i3Node).SetAttribute('border-style', 'dotted');
                 TDOMElement(i3Node).SetAttribute('border-color', 'black');
                 TDOMElement(i3Node).SetAttribute('margin-top', '0pt');
                 TDOMElement(i3Node).SetAttribute('padding', '4pt');

                 i4Node:=xsl.CreateElement('fo:inline-container');
                 i3Node.Appendchild(i4Node);
                 TDOMElement(i4Node).SetAttribute('width', '12%');

                 eNode:=xsl.CreateElement('fo:block');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 TDOMElement(eNode).SetAttribute('text-align', 'center');
                 TDOMElement(eNode).SetAttribute('border-right-width', '2pt');
                 TDOMElement(eNode).SetAttribute('border-right-style', 'solid');
                 TDOMElement(eNode).SetAttribute('border-right-color', 'black');
                 TDOMElement(eNode).SetAttribute('margin-left', '2pt');
                 TDOMElement(eNode).SetAttribute('margin-right', '2pt');

                 tNode:=xsl.CreateTextNode('COD');
                 eNode.Appendchild(tNode);

                 i4Node:=xsl.CreateElement('fo:inline-container');
                 i3Node.Appendchild(i4Node);
                 TDOMElement(i4Node).SetAttribute('width', '52%');

                 eNode:=xsl.CreateElement('fo:block');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 TDOMElement(eNode).SetAttribute('text-align', 'center');
                 TDOMElement(eNode).SetAttribute('border-right-width', '2pt');
                 TDOMElement(eNode).SetAttribute('border-right-style', 'solid');
                 TDOMElement(eNode).SetAttribute('border-right-color', 'black');
                 TDOMElement(eNode).SetAttribute('margin-left', '2pt');
                 TDOMElement(eNode).SetAttribute('margin-right', '2pt');

                 tNode:=xsl.CreateTextNode('DESC');
                 eNode.Appendchild(tNode);

                 i4Node:=xsl.CreateElement('fo:inline-container');
                 i3Node.Appendchild(i4Node);
                 TDOMElement(i4Node).SetAttribute('width', '8%');

                 eNode:=xsl.CreateElement('fo:block');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 TDOMElement(eNode).SetAttribute('text-align', 'center');
                 TDOMElement(eNode).SetAttribute('border-right-width', '2pt');
                 TDOMElement(eNode).SetAttribute('border-right-style', 'solid');
                 TDOMElement(eNode).SetAttribute('border-right-color', 'black');
                 TDOMElement(eNode).SetAttribute('margin-left', '2pt');
                 TDOMElement(eNode).SetAttribute('margin-right', '2pt');

                 tNode:=xsl.CreateTextNode('UNI');
                 eNode.Appendchild(tNode);

                 i4Node:=xsl.CreateElement('fo:inline-container');
                 i3Node.Appendchild(i4Node);
                 TDOMElement(i4Node).SetAttribute('width', '8%');

                 eNode:=xsl.CreateElement('fo:block');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 TDOMElement(eNode).SetAttribute('text-align', 'center');
                 TDOMElement(eNode).SetAttribute('border-right-width', '2pt');
                 TDOMElement(eNode).SetAttribute('border-right-style', 'solid');
                 TDOMElement(eNode).SetAttribute('border-right-color', 'black');
                 TDOMElement(eNode).SetAttribute('margin-left', '2pt');
                 TDOMElement(eNode).SetAttribute('margin-right', '2pt');

                 tNode:=xsl.CreateTextNode('QTDE');
                 eNode.Appendchild(tNode);

                 i4Node:=xsl.CreateElement('fo:inline-container');
                 i3Node.Appendchild(i4Node);
                 TDOMElement(i4Node).SetAttribute('width', '10%');

                 eNode:=xsl.CreateElement('fo:block');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 TDOMElement(eNode).SetAttribute('text-align', 'center');
                 TDOMElement(eNode).SetAttribute('border-right-width', '2pt');
                 TDOMElement(eNode).SetAttribute('border-right-style', 'solid');
                 TDOMElement(eNode).SetAttribute('border-right-color', 'black');
                 TDOMElement(eNode).SetAttribute('margin-left', '2pt');
                 TDOMElement(eNode).SetAttribute('padding-right', '2pt');

                 tNode:=xsl.CreateTextNode('R$ UNIT');
                 eNode.Appendchild(tNode);

                 i4Node:=xsl.CreateElement('fo:inline-container');
                 i3Node.Appendchild(i4Node);
                 TDOMElement(i4Node).SetAttribute('width', '10%');

                 eNode:=xsl.CreateElement('fo:block');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 TDOMElement(eNode).SetAttribute('text-align', 'center');
                 TDOMElement(eNode).SetAttribute('border-right-width', '0pt');
                 TDOMElement(eNode).SetAttribute('margin-left', '2pt');
                 TDOMElement(eNode).SetAttribute('margin-right', '2pt');

                 tNode:=xsl.CreateTextNode('R$ TOT');
                 eNode.Appendchild(tNode);

                 i4Node:=xsl.CreateElement('xsl:for-each');
                 i3Node.Appendchild(i4Node);
                 TDOMElement(i4Node).SetAttribute('select', 'produto');

                 eNode:=xsl.CreateElement('fo:block');
                 i4Node.Appendchild(eNode);

                 i5Node:=xsl.CreateElement('fo:inline-container');
                 i4Node.Appendchild(i5Node);
                 TDOMElement(i5Node).SetAttribute('width', '12%');

                 eNode:=xsl.CreateElement('fo:block');
                 i5Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('text-align', 'left');
                 TDOMElement(eNode).SetAttribute('border-right-width', '2pt');
                 TDOMElement(eNode).SetAttribute('border-right-style', 'solid');
                 TDOMElement(eNode).SetAttribute('border-right-color', 'black');
                 TDOMElement(eNode).SetAttribute('margin-left', '2pt');
                 TDOMElement(eNode).SetAttribute('margin-right', '2pt');

                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'codigo');

                 i5Node:=xsl.CreateElement('fo:inline-container');
                 i4Node.Appendchild(i5Node);
                 TDOMElement(i5Node).SetAttribute('width', '52%');

                 eNode:=xsl.CreateElement('fo:block');
                 i5Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('text-align', 'left');
                 TDOMElement(eNode).SetAttribute('border-right-width', '2pt');
                 TDOMElement(eNode).SetAttribute('border-right-style', 'solid');
                 TDOMElement(eNode).SetAttribute('border-right-color', 'black');
                 TDOMElement(eNode).SetAttribute('margin-left', '2pt');
                 TDOMElement(eNode).SetAttribute('margin-right', '2pt');

                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'descricao');

                 i5Node:=xsl.CreateElement('fo:inline-container');
                 i4Node.Appendchild(i5Node);
                 TDOMElement(i5Node).SetAttribute('width', '8%');

                 eNode:=xsl.CreateElement('fo:block');
                 i5Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('text-align', 'center');
                 TDOMElement(eNode).SetAttribute('border-right-width', '2pt');
                 TDOMElement(eNode).SetAttribute('border-right-style', 'solid');
                 TDOMElement(eNode).SetAttribute('border-right-color', 'black');
                 TDOMElement(eNode).SetAttribute('margin-left', '2pt');
                 TDOMElement(eNode).SetAttribute('margin-right', '2pt');

                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'unidade');

                 i5Node:=xsl.CreateElement('fo:inline-container');
                 i4Node.Appendchild(i5Node);
                 TDOMElement(i5Node).SetAttribute('width', '8%');

                 eNode:=xsl.CreateElement('fo:block');
                 i5Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('text-align', 'center');
                 TDOMElement(eNode).SetAttribute('border-right-width', '2pt');
                 TDOMElement(eNode).SetAttribute('border-right-style', 'solid');
                 TDOMElement(eNode).SetAttribute('border-right-color', 'black');
                 TDOMElement(eNode).SetAttribute('margin-left', '2pt');
                 TDOMElement(eNode).SetAttribute('margin-right', '2pt');

                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'quantidade');

                 i5Node:=xsl.CreateElement('fo:inline-container');
                 i4Node.Appendchild(i5Node);
                 TDOMElement(i5Node).SetAttribute('width', '10%');

                 eNode:=xsl.CreateElement('fo:block');
                 i5Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('text-align', 'right');
                 TDOMElement(eNode).SetAttribute('border-right-width', '2pt');
                 TDOMElement(eNode).SetAttribute('border-right-style', 'solid');
                 TDOMElement(eNode).SetAttribute('border-right-color', 'black');
                 TDOMElement(eNode).SetAttribute('margin-left', '2pt');
                 TDOMElement(eNode).SetAttribute('padding-right', '2pt');

                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'valor_unitario');

                 i5Node:=xsl.CreateElement('fo:inline-container');
                 i4Node.Appendchild(i5Node);
                 TDOMElement(i5Node).SetAttribute('width', '10%');

                 eNode:=xsl.CreateElement('fo:block');
                 i5Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('text-align', 'right');
                 TDOMElement(eNode).SetAttribute('border-right-width', '0pt');
                 TDOMElement(eNode).SetAttribute('margin-left', '2pt');
                 TDOMElement(eNode).SetAttribute('margin-right', '2pt');

                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'valor_total');

                 i3Node:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(i3Node);
                 TDOMElement(i3Node).SetAttribute('font-size', '10pt');
                 TDOMElement(i3Node).SetAttribute('font-family', 'sans-serif');
                 TDOMElement(i3Node).SetAttribute('color', '#303030');
                 TDOMElement(i3Node).SetAttribute('space-after.optimum', '5pt');
                 TDOMElement(i3Node).SetAttribute('text-align', 'left');
                 TDOMElement(i3Node).SetAttribute('padding', '4pt');
                 TDOMElement(i3Node).SetAttribute('margin-top', '15pt');
                 TDOMElement(i3Node).SetAttribute('border-width', '3pt');
                 TDOMElement(i3Node).SetAttribute('border-style', 'dotted');
                 TDOMElement(i3Node).SetAttribute('border-color', 'black');

                 i4Node:=xsl.CreateElement('fo:inline');
                 i3Node.Appendchild(i4Node);

                 tNode:=xsl.CreateTextNode('Valor total dos produtos: ');
                 i4Node.Appendchild(tNode);
                 eNode:=xsl.CreateElement('xsl:value-of');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('select', 'total_produtos');
                 tNode:=xsl.CreateElement('fo:block');
                 i4Node.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode('Valor total da nota: ');
                 i4Node.Appendchild(tNode);
                 eNode:=xsl.CreateElement('xsl:value-of');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('select', 'total_nota');

                 eNode:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-size', '0pt');
                 TDOMElement(eNode).SetAttribute('margin-top', '15pt');
                 TDOMElement(eNode).SetAttribute('border-bottom-width', '3pt');
                 TDOMElement(eNode).SetAttribute('border-bottom-style', 'double');
                 TDOMElement(eNode).SetAttribute('border-bottom-color', '#990000');

                 eNode:=xsl.CreateElement('fo:block');
                 iNode.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-size', '0pt');
                 TDOMElement(eNode).SetAttribute('margin-top', '15pt');
                 TDOMElement(eNode).SetAttribute('border-bottom-width', '3pt');
                 TDOMElement(eNode).SetAttribute('border-bottom-style', 'double');
                 TDOMElement(eNode).SetAttribute('border-bottom-color', '#990000');

                 eNode:=xsl.CreateElement('fo:block');
                 iNode.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-size', '0pt');
                 TDOMElement(eNode).SetAttribute('margin-top', '15pt');
                 TDOMElement(eNode).SetAttribute('border-bottom-width', '3pt');
                 TDOMElement(eNode).SetAttribute('border-bottom-style', 'double');
                 TDOMElement(eNode).SetAttribute('border-bottom-color', '#990000');

                 i3Node:=xsl.CreateElement('fo:block');
                 iNode.Appendchild(i3Node);
                 TDOMElement(i3Node).SetAttribute('font-size', '10pt');
                 TDOMElement(i3Node).SetAttribute('font-family', 'sans-serif');
                 TDOMElement(i3Node).SetAttribute('color', '#303030');
                 TDOMElement(i3Node).SetAttribute('space-after.optimum', '5pt');
                 TDOMElement(i3Node).SetAttribute('text-align', 'left');
                 TDOMElement(i3Node).SetAttribute('padding', '4pt');
                 TDOMElement(i3Node).SetAttribute('margin-top', '15pt');
                 TDOMElement(i3Node).SetAttribute('border-width', '3pt');
                 TDOMElement(i3Node).SetAttribute('border-style', 'dotted');
                 TDOMElement(i3Node).SetAttribute('border-color', 'black');

                 i4Node:=xsl.CreateElement('fo:inline');
                 i3Node.Appendchild(i4Node);

                 tNode:=xsl.CreateTextNode('Quantidade de produtos: ');
                 i4Node.Appendchild(tNode);
                 eNode:=xsl.CreateElement('xsl:value-of');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('select', 'quantidade_produtos');
                 tNode:=xsl.CreateElement('fo:block');
                 i4Node.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode('Valor total dos produtos: ');
                 i4Node.Appendchild(tNode);
                 eNode:=xsl.CreateElement('xsl:value-of');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('select', 'total_produtos_relatorio');
                 tNode:=xsl.CreateElement('fo:block');
                 i4Node.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode('Quantidade de notas: ');
                 i4Node.Appendchild(tNode);
                 eNode:=xsl.CreateElement('xsl:value-of');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('select', 'quantidade_notas');
                 tNode:=xsl.CreateElement('fo:block');
                 i4Node.Appendchild(tNode);
                 tNode:=xsl.CreateTextNode('Valor total das notas: ');
                 i4Node.Appendchild(tNode);
                 eNode:=xsl.CreateElement('xsl:value-of');
                 i4Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('select', 'total_relatorio');

                 iNode:=xsl.CreateElement('xsl:template');
                 rNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('match', 'emitente');

                 i2Node:=xsl.CreateElement('fo:block');
                 iNode.Appendchild(i2Node);
                 TDOMElement(i2Node).SetAttribute('font-size', '10pt');
                 TDOMElement(i2Node).SetAttribute('font-family', 'sans-serif');
                 TDOMElement(i2Node).SetAttribute('color', '#303030');
                 TDOMElement(i2Node).SetAttribute('space-after.optimum', '5pt');
                 TDOMElement(i2Node).SetAttribute('line-height', '14pt');
                 TDOMElement(i2Node).SetAttribute('text-align', 'left');
                 TDOMElement(i2Node).SetAttribute('padding', '4pt');
                 TDOMElement(i2Node).SetAttribute('margin-top', '0pt');
                 TDOMElement(i2Node).SetAttribute('border-width', '3pt');
                 TDOMElement(i2Node).SetAttribute('border-style', 'dotted');
                 TDOMElement(i2Node).SetAttribute('border-color', 'black');

                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 tNode:=xsl.CreateTextNode('Nome / Razao Social: ');
                 eNode.Appendchild(tNode);
                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'nome');
                 tNode:=xsl.CreateTextNode(' ');
                 eNode.Appendchild(tNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'fantasia');
                 eNode:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(eNode);

                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 tNode:=xsl.CreateTextNode('CPF / CNPJ: ');
                 eNode.Appendchild(tNode);
                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'documento');
                 eNode:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(eNode);

                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 tNode:=xsl.CreateTextNode('Endereco: ');
                 eNode.Appendchild(tNode);
                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'endereco');
                 eNode:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(eNode);

                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 tNode:=xsl.CreateTextNode('Cidade: ');
                 eNode.Appendchild(tNode);
                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'cidade');
                 eNode:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(eNode);

                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 tNode:=xsl.CreateTextNode('E-mail: ');
                 eNode.Appendchild(tNode);
                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'email');
                 eNode:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(eNode);

                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 tNode:=xsl.CreateTextNode('Telefone: ');
                 eNode.Appendchild(tNode);
                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'telefone');
                 eNode:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(eNode);


                 iNode:=xsl.CreateElement('xsl:template');
                 rNode.Appendchild(iNode);
                 TDOMElement(iNode).SetAttribute('match', 'destinatario');

                 i2Node:=xsl.CreateElement('fo:block');
                 iNode.Appendchild(i2Node);
                 TDOMElement(i2Node).SetAttribute('font-size', '10pt');
                 TDOMElement(i2Node).SetAttribute('font-family', 'sans-serif');
                 TDOMElement(i2Node).SetAttribute('color', '#303030');
                 TDOMElement(i2Node).SetAttribute('space-after.optimum', '5pt');
                 TDOMElement(i2Node).SetAttribute('line-height', '14pt');
                 TDOMElement(i2Node).SetAttribute('text-align', 'left');
                 TDOMElement(i2Node).SetAttribute('padding', '4pt');
                 TDOMElement(i2Node).SetAttribute('margin-top', '0pt');
                 TDOMElement(i2Node).SetAttribute('border-width', '3pt');
                 TDOMElement(i2Node).SetAttribute('border-style', 'dotted');
                 TDOMElement(i2Node).SetAttribute('border-color', 'black');

                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 tNode:=xsl.CreateTextNode('Nome / Razao Social: ');
                 eNode.Appendchild(tNode);
                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'nome');
                 eNode:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(eNode);

                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 tNode:=xsl.CreateTextNode('CPF / CNPJ: ');
                 eNode.Appendchild(tNode);
                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'documento');
                 eNode:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(eNode);

                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 tNode:=xsl.CreateTextNode('Endereco: ');
                 eNode.Appendchild(tNode);
                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'endereco');
                 eNode:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(eNode);

                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 tNode:=xsl.CreateTextNode('Cidade: ');
                 eNode.Appendchild(tNode);
                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'cidade');
                 eNode:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(eNode);

                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 tNode:=xsl.CreateTextNode('E-mail: ');
                 eNode.Appendchild(tNode);
                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'email');
                 eNode:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(eNode);

                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 TDOMElement(eNode).SetAttribute('font-weight', 'bold');
                 tNode:=xsl.CreateTextNode('Telefone: ');
                 eNode.Appendchild(tNode);
                 eNode:=xsl.CreateElement('fo:inline');
                 i2Node.Appendchild(eNode);
                 tNode:=xsl.CreateElement('xsl:value-of');
                 eNode.Appendchild(tNode);
                 TDOMElement(tNode).SetAttribute('select', 'telefone');
                 eNode:=xsl.CreateElement('fo:block');
                 i2Node.Appendchild(eNode);

                 if (tipo=2) then
                 begin
                     iNode:=xsl.CreateElement('xsl:template');
                     rNode.Appendchild(iNode);
                     TDOMElement(iNode).SetAttribute('match', 'periodo');

                     i2Node:=xsl.CreateElement('fo:block');
                     iNode.Appendchild(i2Node);
                     TDOMElement(i2Node).SetAttribute('font-size', '12pt');
                     TDOMElement(i2Node).SetAttribute('font-family', 'sans-serif');
                     TDOMElement(i2Node).SetAttribute('color', '#990000');
                     TDOMElement(i2Node).SetAttribute('space-after.optimum', '5pt');
                     TDOMElement(i2Node).SetAttribute('text-align', 'left');

                     eNode:=xsl.CreateElement('fo:inline');
                     i2Node.Appendchild(eNode);
                     tNode:=xsl.CreateTextNode('Notas emitidas entre ');
                     eNode.Appendchild(tNode);
                     tNode:=xsl.CreateElement('xsl:value-of');
                     eNode.Appendchild(tNode);
                     TDOMElement(tNode).SetAttribute('select', 'inicio');
                     tNode:=xsl.CreateTextNode(' e ');
                     eNode.Appendchild(tNode);
                     tNode:=xsl.CreateElement('xsl:value-of');
                     eNode.Appendchild(tNode);
                     TDOMElement(tNode).SetAttribute('select', 'fim');
                 end;

                 //FIM XSL

           end;

           if not DirectoryExists(caminho) then
           begin
                try
                   CreateDir (caminho);
                except
                   Screen.Cursor:=crDefault;
                   Application.MessageBox('Falha ao criar diretório. Tente novamente','Erro',0);
                   Exit;
                end;
           end;

           writeXMLFile(xml,caminho+'\relatorio.xml');
           if (TipoRel.ItemIndex>3) then
              writeXMLFile(xsd,caminho+'\relatorio.xsd');
           writeXMLFile(xsl,caminho+'\relatorio.xsl');
           if (TipoRel.ItemIndex<=3) then
           begin
               WinExec(PChar('cmd /k cd '+Application.Location+'\fop'+
                                  '&& cmd /k fop -xml '+caminho+'\relatorio.xml -xsl '+caminho+
                                  '\relatorio.xsl -pdf '+caminho+'\relatorio.pdf'), SW_HIDE);
           end;
           Application.MessageBox('Relatório criado com sucesso.','Relatório gerado',0);
        except
          Application.MessageBox('Erro ao criar o relatório desejado.','Erro',0);
        end;
      xml.Free;
      if (TipoRel.ItemIndex>3) then
         xsd.Free;
      xsl.Free;
    end;
    Screen.Cursor:=crDefault;
    BtnCancelar.SetFocus;
end;


procedure TFRel.FormShow(Sender: TObject);
begin
  QDest.Close;
  QDest.SQL.Text:='select des_nom from destinatarios inner join notas on not_dest=des_doc where not_usr='+inttostr(user)+' group by des_nom order by des_nom';
  QEmit.Close;
  QEmit.SQL.Text:='select emi_rzo from emitentes inner join notas on not_emit=emi_doc where not_usr='+inttostr(user)+' group by emi_rzo order by emi_rzo';
  try
    QDest.Open;
    QEmit.Open;
  except
    ShowMessage('Problema com a conexão com o banco de dados');
    Close;
  end;
  ComboDest.Items.Clear;
  ComboEmit.Items.Clear;
  with QDest do
  begin
    if RecNo>0 then
    begin
         First;
         while not (EOF) do
         begin
           ComboDest.Items.Add(FieldByName('des_nom').AsString);
           Next;
         end;
    end;
  end;
  with QEmit do
  begin
    if RecNo>0 then
    begin
         First;
         while not (EOF) do
         begin
           ComboEmit.Items.Add(FieldByName('emi_rzo').AsString);
           Next;
         end;
    end;
  end;
  ComboDest.ItemIndex:=0;
  ComboEmit.ItemIndex:=0;
  DataIni.Date:=IncYear(Today,-1);
  DataFim.Date:=Today;
end;



procedure TFRel.TipoRelClick(Sender: TObject);
begin
  if (TipoRel.ItemIndex=0) or (TipoRel.ItemIndex=4) then //geral
  begin
    LblNome.Visible:=false;
    ComboEmit.Visible:=false;
    ComboDest.Visible:=false;
    Periodo.Visible:=false;
  end
  else if (TipoRel.ItemIndex=1) or (TipoRel.ItemIndex=5) then //por periodo
  begin
    LblNome.Caption:='Período:';
    LblNome.Visible:=true;
    ComboEmit.Visible:=false;
    ComboDest.Visible:=false;
    Periodo.Visible:=true;
  end
  else if (TipoRel.ItemIndex=2) or (TipoRel.ItemIndex=6) then //por emitente
  begin
    LblNome.Visible:=true;
    LblNome.Caption:='Emitente:';
    ComboEmit.Visible:=true;
    ComboDest.Visible:=false;
    Periodo.Visible:=false;
  end
  else if (TipoRel.ItemIndex=3) or (TipoRel.ItemIndex=7) then //por destinatario
  begin
    LblNome.Visible:=true;
    LblNome.Caption:='Destinatário:';
    ComboEmit.Visible:=false;
    ComboDest.Visible:=true;
    Periodo.Visible:=false;
  end;
end;

end.

