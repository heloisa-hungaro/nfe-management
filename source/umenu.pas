unit UMenu;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, mysql40conn, mysql50conn, sqldb, FileUtil, Forms,
  Controls, Graphics, Dialogs, Menus, StdCtrls, ExtCtrls, IniFiles,
  ComCtrls, Buttons, laz2_XMLRead, laz2_DOM;

type

  { TFMenu }

  TFMenu = class(TForm)
    BtnAlterar: TButton;
    BtnVoltar: TButton;
    BtnEntrar: TButton;
    BtnCancelar: TButton;
    BtnCadastrar: TButton;
    BtnVoltar1: TButton;
    MostraSenhaNova1: TCheckBox;
    Servidor: TPanel;
    EdtAltSenha: TEdit;
    EdtAltUser: TEdit;
    EdtPorta: TEdit;
    EdtServ: TEdit;
    Label10: TLabel;
    Label11: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    MostraSenhaNova: TCheckBox;
    EdtSenha: TEdit;
    EdtNovaSenha: TEdit;
    EdtUser: TEdit;
    EdtNovoUser: TEdit;
    ItemRemove: TMenuItem;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Cadastro: TPanel;
    ItemTroca: TMenuItem;
    MenuItens: TMainMenu;
    Importa: TMenuItem;
    ItemImporta: TMenuItem;
    ItemRelatorios: TMenuItem;
    ItemVisualiza: TMenuItem;
    ItemSair: TMenuItem;
    Arquivo: TOpenDialog;
    Database: TMySQL50Connection;
    Login: TPanel;
    MostraSenha: TCheckBox;
    Botoes: TPanel;
    AlteraServ: TPanel;
    Panel2: TPanel;
    NovoUser: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Q: TSQLQuery;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    SpeedButton5: TSpeedButton;
    SpeedButton6: TSpeedButton;
    StatusBar: TStatusBar;
    Transaction: TSQLTransaction;
    procedure AlteraServClick(Sender: TObject);
    procedure AlteraServMouseEnter(Sender: TObject);
    procedure AlteraServMouseLeave(Sender: TObject);
    procedure BtnAlterarClick(Sender: TObject);
    procedure BtnCadastrarClick(Sender: TObject);
    procedure BtnCancelarClick(Sender: TObject);
    procedure BtnEntrarClick(Sender: TObject);
    procedure BtnVoltar1Click(Sender: TObject);
    procedure BtnVoltarClick(Sender: TObject);
    procedure EdtNovaSenhaChange(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: char);
    procedure ItemRelatoriosClick(Sender: TObject);
    procedure ItemRemoveClick(Sender: TObject);
    procedure ItemVisualizaClick(Sender: TObject);
    procedure MostraSenhaChange(Sender: TObject);
    procedure MostraSenhaNova1Change(Sender: TObject);
    procedure MostraSenhaNovaChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ItemImportaClick(Sender: TObject);
    procedure ItemSairClick(Sender: TObject);
    procedure ItemTrocaClick(Sender: TObject);
    procedure NovoUserClick(Sender: TObject);
    procedure NovoUserMouseEnter(Sender: TObject);
    procedure NovoUserMouseLeave(Sender: TObject);
  private
    { private declarations }
  public
        user: integer;
    { public declarations }
  end;

var
  FMenu: TFMenu;

implementation

uses URel, URemove, UVisualiza;
{$R *.lfm}

{ TFMenu }


procedure TFMenu.ItemSairClick(Sender: TObject);
begin
  Close;
end;

procedure TFMenu.ItemTrocaClick(Sender: TObject);
begin
  MenuItens.Items[0].Visible:=false;
  MenuItens.Items[1].Visible:=false;
  MenuItens.Items[2].Visible:=false;
  MenuItens.Items[3].Visible:=false;
  MenuItens.Items[4].Visible:=false;
  MenuItens.Items[5].Visible:=false;
  StatusBar.Visible:=false;
  Login.Visible:=true;
  Botoes.Visible:=false;
  EdtUser.SetFocus;
end;

procedure TFMenu.NovoUserClick(Sender: TObject);
begin
  Cadastro.Visible:=true;
  EdtNovoUser.SetFocus;
  Login.Visible:=false;
end;

procedure TFMenu.NovoUserMouseEnter(Sender: TObject);
begin
     Screen.Cursor:=crHandPoint;
     NovoUser.Font.Underline:=True;
end;

procedure TFMenu.NovoUserMouseLeave(Sender: TObject);
begin
     NovoUser.Font.Underline:=False;
     Screen.Cursor:=crDefault;
end;


function DadosEmit(node: TDomNode; Q:TSQLQuery; data: string): String;
var inode: TDOMNode; doc: string;
begin
       node:=node.FindNode('emit');
       inode:=node.FindNode('CNPJ');
       if inode <> nil then
          doc:=inode.FirstChild.NodeValue
       else
       begin
            inode:=node.FindNode('CPF');
            doc:=inode.FirstChild.NodeValue;
       end;

       Q.Close;
       Q.SQL.Clear;
       Q.SQL.Text:='select emi_doc, emi_dat from emitentes where emi_doc='+doc;
       Q.Open;
       if (Q.RecNo>0) then
       begin
           if (FormatDateTime('yyyy-mm-dd',Q.FieldByName('emi_dat').AsDateTime)<data) then
           begin
                Q.SQL.Clear;
                Q.SQL.Text:='update emitentes set emi_rzo=:RZO, emi_fan=:FAN, emi_end=:END, '+
                               'emi_cid=:CID, emi_tel=:TEL, emi_eml=:EML, emi_dat='+QuotedStr(data)+
                               ' where emi_doc=:DOC';
           end
           else
           begin
               DadosEmit:=doc;
               Exit;
           end;
       end
       else
       begin
            Q.Close;
            Q.SQL.Clear;
            Q.SQL.Text:='insert into emitentes (emi_doc, emi_rzo, emi_fan, emi_end, '+
                           'emi_cid, emi_tel, emi_eml, emi_dat) values (:DOC,:RZO,:FAN,'+
                           ':END,:CID,:TEL,:EML,'+QuotedStr(data)+')';
       end;

       Q.Params.ParamByName('DOC').AsString:=doc;
       inode:=node.FindNode('xNome');
       if inode <> nil then
          Q.Params.ParamByName('RZO').AsString:=AnsiUpperCase(inode.FirstChild.NodeValue);
       inode:=node.FindNode('xFant');
       if inode <> nil then
          Q.Params.ParamByName('FAN').AsString:=AnsiUpperCase(inode.FirstChild.NodeValue);
       node:=node.FindNode('enderEmit');
       inode:=node.FindNode('xLgr');
       if inode <> nil then
          Q.Params.ParamByName('END').AsString:=AnsiUpperCase(inode.FirstChild.NodeValue);
       inode:=node.FindNode('nro');
       if inode <> nil then
          Q.Params.ParamByName('END').AsString:=Q.Params.ParamByName('END').AsString+' '+AnsiUpperCase(inode.FirstChild.NodeValue);
       inode:=node.FindNode('xCpl');
       if inode <> nil then
          Q.Params.ParamByName('END').AsString:=Q.Params.ParamByName('END').AsString+' ('+AnsiUpperCase(inode.FirstChild.NodeValue)+')';
       inode:=node.FindNode('xBairro');
       if inode <> nil then
          Q.Params.ParamByName('END').AsString:=Q.Params.ParamByName('END').AsString+', '+AnsiUpperCase(inode.FirstChild.NodeValue);
       inode:=node.FindNode('CEP');
       if inode <> nil then
          Q.Params.ParamByName('END').AsString:=Q.Params.ParamByName('END').AsString+' - CEP '+inode.FirstChild.NodeValue;
       inode:=node.FindNode('xMun');
       if inode <> nil then
          Q.Params.ParamByName('CID').AsString:=AnsiUpperCase(inode.FirstChild.NodeValue);
       inode:=node.FindNode('UF');
       if inode <> nil then
          Q.Params.ParamByName('CID').AsString:=Q.Params.ParamByName('CID').AsString+'/'+AnsiUpperCase(inode.FirstChild.NodeValue);
       inode:=node.FindNode('xPais');
       if inode <> nil then
          Q.Params.ParamByName('CID').AsString:=Q.Params.ParamByName('CID').AsString+' - '+AnsiUpperCase(inode.FirstChild.NodeValue);

       inode:=node.FindNode('fone');
       if inode <> nil then
          Q.Params.ParamByName('TEL').AsString:=inode.FirstChild.NodeValue;
       inode:=node.FindNode('Email');
       if inode <> nil then
          Q.Params.ParamByName('EML').AsString:=AnsiLowerCase(inode.FirstChild.NodeValue);
       Q.ExecSQL;
       DadosEmit:=doc;
end;

function DadosDest(node: TDomNode; Q:TSQLQuery; data: string): String;
var inode: TDOMNode; doc: string;
begin
       node:=node.FindNode('dest');
       inode:=node.FindNode('CNPJ');
       if inode <> nil then
          doc:=inode.FirstChild.NodeValue
       else
       begin
            inode:=node.FindNode('CPF');
            doc:=inode.FirstChild.NodeValue;
       end;

       Q.Close;
       Q.SQL.Clear;
       Q.SQL.Text:='select des_doc, des_dat from destinatarios where des_doc='+doc;
       Q.Open;

       if (Q.RecNo>0) then
       begin
            if (FormatDateTime('yyyy-mm-dd',Q.FieldByName('des_dat').AsDateTime)<data) then
            begin
                Q.SQL.Clear;
                Q.SQL.Text:='update destinatarios set des_nom=:NOM, des_end=:END, des_cid=:CID, '+
                               'des_tel=:TEL, des_eml=:EML, des_dat='+QuotedStr(data)+
                               ' where des_doc=:DOC';
            end
            else
            begin
                 DadosDest:=doc;
                 Exit;
            end;
       end
       else
       begin
            Q.Close;
            Q.SQL.Clear;
            Q.SQL.Text:='insert into destinatarios (des_doc, des_nom, des_end, '+
                           'des_cid, des_tel, des_eml, des_dat) values (:DOC,:NOM,'+
                           ':END,:CID,:TEL,:EML,'+QuotedStr(data)+')';
       end;

       Q.Params.ParamByName('DOC').AsString:=doc;
       inode:=node.FindNode('xNome');
       if inode <> nil then
          Q.Params.ParamByName('NOM').AsString:=AnsiUpperCase(inode.FirstChild.NodeValue);
       inode:=node.FindNode('email');
       if inode <> nil then
          Q.Params.ParamByName('EML').AsString:=AnsiLowerCase(inode.FirstChild.NodeValue);
       node:=node.FindNode('enderDest');
       inode:=node.FindNode('xLgr');
       if inode <> nil then
          Q.Params.ParamByName('END').AsString:=AnsiUpperCase(inode.FirstChild.NodeValue);
       inode:=node.FindNode('nro');
       if inode <> nil then
          Q.Params.ParamByName('END').AsString:=Q.Params.ParamByName('END').AsString+' '+AnsiUpperCase(inode.FirstChild.NodeValue);
       inode:=node.FindNode('xCpl');
       if inode <> nil then
          Q.Params.ParamByName('END').AsString:=Q.Params.ParamByName('END').AsString+' ('+AnsiUpperCase(inode.FirstChild.NodeValue)+')';
       inode:=node.FindNode('xBairro');
       if inode <> nil then
          Q.Params.ParamByName('END').AsString:=Q.Params.ParamByName('END').AsString+', '+AnsiUpperCase(inode.FirstChild.NodeValue);
       inode:=node.FindNode('CEP');
       if inode <> nil then
          Q.Params.ParamByName('END').AsString:=Q.Params.ParamByName('END').AsString+' - CEP '+inode.FirstChild.NodeValue;
       inode:=node.FindNode('xMun');
       if inode <> nil then
          Q.Params.ParamByName('CID').AsString:=AnsiUpperCase(inode.FirstChild.NodeValue);
       inode:=node.FindNode('UF');
       if inode <> nil then
          Q.Params.ParamByName('CID').AsString:=Q.Params.ParamByName('CID').AsString+'/'+AnsiUpperCase(inode.FirstChild.NodeValue);
       inode:=node.FindNode('xPais');
       if inode <> nil then
          Q.Params.ParamByName('CID').AsString:=Q.Params.ParamByName('CID').AsString+' - '+AnsiUpperCase(inode.FirstChild.NodeValue);
       inode:=node.FindNode('fone');
       if inode <> nil then
          Q.Params.ParamByName('TEL').AsString:=inode.FirstChild.NodeValue;
       Q.ExecSQL;
       DadosDest:=doc;
end;

procedure DadosNota(node: TDomNode; Q:TSQLQuery; data: string; emit: string; dest: string; chave: string; user: string);
var inode: TDOMNode;
begin
       node:=(node.FindNode('total')).FindNode('ICMSTot');
       Q.Close;
       Q.SQL.Clear;
       Q.SQL.Text:='insert into notas (not_usr, not_cha, not_emit, not_dest, not_emi, '+
                           'not_imp, not_pro, not_tot) values ('+user+','+QuotedStr(chave)+','+
                           emit+','+dest+','+QuotedStr(data)+','+QuotedStr(FormatDateTime('yyyy-mm-dd', Date))+',:PRO,:TOT)';

       inode:=node.FindNode('vProd');
       if inode <> nil then
          Q.Params.ParamByName('PRO').AsString:=inode.FirstChild.NodeValue;
       inode:=node.FindNode('vNF');
       if inode <> nil then
          Q.Params.ParamByName('TOT').AsString:=inode.FirstChild.NodeValue;
       Q.ExecSQL;
end;

procedure DadosProdutos(node: TDomNode; Q:TSQLQuery; chave: string);
var inode: TDOMNode;
begin
       node:=node.FindNode('prod');
       Q.Close;
       Q.SQL.Clear;
       Q.SQL.Text:='insert into produtos (pro_not, pro_cod, pro_des, pro_uni, '+
                           'pro_qtd, pro_unt, pro_tot) values ('+QuotedStr(chave)+',:COD,:DES,'+
                           ':UNI,:QTD,:UNT,:TOT)';

       inode:=node.FindNode('cProd');
       if inode <> nil then
          Q.Params.ParamByName('COD').AsString:=inode.FirstChild.NodeValue;
       inode:=node.FindNode('xProd');
       if inode <> nil then
          Q.Params.ParamByName('DES').AsString:=inode.FirstChild.NodeValue;
       inode:=node.FindNode('uCom');
       if inode <> nil then
          Q.Params.ParamByName('UNI').AsString:=AnsiUpperCase(inode.FirstChild.NodeValue);
       inode:=node.FindNode('qCom');
       if inode <> nil then
          Q.Params.ParamByName('QTD').AsString:=inode.FirstChild.NodeValue;
       inode:=node.FindNode('vUnCom');
       if inode <> nil then
          Q.Params.ParamByName('UNT').AsString:=inode.FirstChild.NodeValue;
       inode:=node.FindNode('vProd');
       if inode <> nil then
          Q.Params.ParamByName('TOT').AsString:=inode.FirstChild.NodeValue;
       Q.ExecSQL;
end;

procedure TFMenu.ItemImportaClick(Sender: TObject);
var nomearq : string; node: TDOMNode; doc: TXMLDocument;  data,emit,dest,chave: string;
begin
  Screen.Cursor:=crHourGlass;
  if Arquivo.Execute then
  begin
    try
       nomearq := Arquivo.Filename;
       ReadXMLFile(doc, nomearq);
       if AnsiLowerCase(doc.DocumentElement.NodeName)<>'nfeproc' then
       begin
          Application.MessageBox('Não é arquivo XML de NF-e','Arquivo inválido',0);
          Screen.Cursor:=crDefault;
          Exit;
       end;

       node:=doc.DocumentElement.FindNode('protNFe');
       if node.Attributes.GetNamedItem('versao').NodeValue<>'3.10' then
       begin
          Application.MessageBox('Versão de layout da NF-e deve ser 3.10 (mais atual em jan/2016)','Layout de arquivo de versão inválida',0);
          Screen.Cursor:=crDefault;
          Exit;
       end;
       node:=(node.FindNode('infProt')).FindNode('chNFe');
       if node <> nil then
          chave:=node.FirstChild.NodeValue;

       Q.Close;
       Q.SQL.Clear;
       Q.SQL.Text:='select not_cha from notas where not_cha='+QuotedStr(chave)+' and not_usr='+inttostr(user);
       Q.Open;
       if Q.RecNo>0 then
       begin
            Application.MessageBox('Arquivo já importado anteriormente!','Nota já existe no sistema',0);
            Screen.Cursor:=crDefault;
            Exit;
       end;

       node:=(doc.DocumentElement.FindNode('NFe')).FindNode('infNFe');
       node:=(node.FindNode('ide')).FindNode('dhEmi');
       data:=copy(node.FirstChild.NodeValue,1,10);
       node:=(node.ParentNode).ParentNode;
       emit:=DadosEmit(node,Q,data);
       dest:=DadosDest(node,Q,data);
       DadosNota(node,Q,data,emit,dest,chave,inttostr(user));
       Q.Close;
       Q.SQL.Clear;

       Q.SQL.Text:='select count(not_cha) as total from notas where not_cha='+QuotedStr(chave);
       Q.Open;

       if (Q.Fieldbyname('total').AsInteger=1) then
       begin
           node:=node.FindNode('det');
           while (node <> nil) and (node.NodeName='det') do
           begin
               DadosProdutos(node,Q,chave);
               node:=node.NextSibling;
           end;
       end;
       Transaction.Commit;
       Application.MessageBox('Importação realizada com sucesso.','Dados importados',0);
    except
       Application.MessageBox('Erro ao processar o arquivo selecionado.','Arquivo inválido',0);
    end;
    doc.Free;
  end;
  Screen.Cursor:=crDefault;
end;

procedure TFMenu.BtnCancelarClick(Sender: TObject);
begin
  Close;
end;

procedure TFMenu.BtnEntrarClick(Sender: TObject);
var
   ini: TINIFile; serv: string;
begin
     Screen.Cursor:=crHourGlass;
     if not (Database.Connected) then
     begin
         try
               ini:=TINIFile.Create('ControleNFe.ini');
               serv:=ini.ReadString('conexao','host','prodec.noip.us');
               Database.HostName:=serv;
               Database.Port:=ini.ReadInteger('conexao','porta',3306);
               Database.UserName:=ini.ReadString('conexao','user','heloisa');
               Database.Password:=ini.ReadString('conexao','psw','123');
               ini.Free;
               Database.Connected:=true;
         except
               Application.MessageBox(PChar('Não foi possível conectar no servidor "'+serv+'" com os dados indicados.'),'Conexão falhou.',0);
               EdtUser.SetFocus;
               Screen.Cursor:=crDefault;
               Exit;
         end;
     end;

   if trim(EdtUser.Text)='' then
   begin
     Application.MessageBox('Digite seu nome de usuário!','Campo usuário vazio',0);
     EdtUser.SetFocus;
     Screen.Cursor:=crDefault;
     Exit;
   end
   else if EdtSenha.Text='' then
   begin
     Application.MessageBox('Digite sua senha!','Campo senha vazio',0);
     EdtSenha.SetFocus;
     Screen.Cursor:=crDefault;
     Exit;
   end;
   Q.Close;
   Q.SQL.Clear;
   Q.SQL.Text:='select * from users where lower(usr_log)='+QuotedStr(AnsiLowerCase(trim(EdtUser.Text)));
   Q.Open;
   if Q.RecNo=0 then
   begin
     Application.MessageBox('Usuário inválido.','Nome de usuário inexistente',0);
     EdtUser.SetFocus;
     Screen.Cursor:=crDefault;
     Exit;
   end;
   Q.Close;
   Q.SQL.Clear;
   Q.SQL.Text:='select * from users where lower(usr_log)='+QuotedStr(AnsiLowerCase(trim(EdtUser.Text)))+' and usr_psw=md5('+QuotedStr(EdtSenha.Text)+')';
   Q.Open;
   if Q.RecNo=0 then
   begin
     Application.MessageBox('Senha inválida.','Campo senha incorreto',0);
     EdtSenha.SetFocus;
     Screen.Cursor:=crDefault;
     Exit;
   end;
   user:=Q.FieldByName('usr_cod').AsInteger;
   MenuItens.Items[0].Visible:=true;
   MenuItens.Items[1].Visible:=true;
   MenuItens.Items[2].Visible:=true;
   MenuItens.Items[3].Visible:=true;
   MenuItens.Items[4].Visible:=true;
   MenuItens.Items[5].Visible:=true;
   StatusBar.Visible:=true;
   StatusBar.Panels[0].Text:=' USUÁRIO: '+AnsiUpperCase(trim(EdtUser.Text));
   StatusBar.Panels[1].Text:=' '+FormatDateTime ('dddd", "dd" de "mmmm" de "yyyy',Now);
   Login.Visible:=false;
   Botoes.Visible:=true;
   EdtSenha.Text:='';
   EdtUser.Text:='';
   Screen.Cursor:=crDefault;
end;

procedure TFMenu.BtnVoltar1Click(Sender: TObject);
begin
     EdtAltSenha.Text:='';
     EdtAltUser.Text:='';
     EdtServ.Text:='';
     EdtPorta.Text:='3306';
     EdtSenha.Text:='';
     EdtUser.Text:='';
     Servidor.Visible:=false;
     Login.Visible:=true;
     EdtUser.SetFocus;
end;

procedure TFMenu.BtnCadastrarClick(Sender: TObject);
begin
  Screen.Cursor:=crHourGlass;
  if trim(EdtNovoUser.Text)='' then
  begin
    Application.MessageBox('Digite um nome de usuário!','Campo usuário vazio',0);
    EdtNovoUser.SetFocus;
    Screen.Cursor:=crDefault;
    Exit;
  end
  else if EdtNovaSenha.Text='' then
  begin
    Application.MessageBox('É necessário escolher uma senha!','Campo senha vazio',0);
    EdtNovaSenha.SetFocus;
    Screen.Cursor:=crDefault;
    Exit;
  end;
  Q.Close;
  Q.SQL.Clear;
  Q.SQL.Text:='select * from users where lower(usr_log)='+QuotedStr(AnsiLowerCase(trim(EdtNovoUser.Text)));
  Q.Open;
  if Q.RecNo>0 then
  begin
    Application.MessageBox('Usuário já existente! Escolha outro.','Nome de usuário indisponível',0);
    EdtNovoUser.SetFocus;
    Screen.Cursor:=crDefault;
    Exit;
  end;
  Q.Close;
  Q.SQL.Clear;
  Q.SQL.Text:='insert into users (usr_log, usr_psw) values (:LOG,md5(:PSW))';
  Q.Params.ParamByName('LOG').AsString:=AnsiLowerCase(trim(EdtNovoUser.Text));
  Q.Params.ParamByName('PSW').AsString:=EdtNovaSenha.Text;
  Q.ExecSQL;
  Transaction.Commit;
  Application.MessageBox('Usuário cadastrado com sucesso.','Cadastro efetuado',0);
  EdtNovaSenha.Text:='';
  EdtNovoUser.Text:='';
  EdtSenha.Text:='';
  EdtUser.Text:='';
  Cadastro.Visible:=false;
  Login.Visible:=true;
  EdtUser.SetFocus;
  Screen.Cursor:=crDefault;
end;

procedure TFMenu.AlteraServMouseEnter(Sender: TObject);
begin
     Screen.Cursor:=crHandPoint;
     AlteraServ.Font.Underline:=True;
end;

procedure TFMenu.AlteraServClick(Sender: TObject);
begin
     Servidor.Visible:=true;
     EdtServ.SetFocus;
     Login.Visible:=false;
end;

procedure TFMenu.AlteraServMouseLeave(Sender: TObject);
begin
     AlteraServ.Font.Underline:=False;
     Screen.Cursor:=crDefault;
end;

procedure TFMenu.BtnAlterarClick(Sender: TObject);
var
   ini: TINIFile;
begin
    Screen.Cursor:=crHourGlass;
    if trim(EdtServ.Text)='' then
    begin
      Application.MessageBox('Digite o servidor!','Campo servidor vazio',0);
      EdtServ.SetFocus;
      Screen.Cursor:=crDefault;
      Exit;
    end else if trim(EdtPorta.Text)='' then
    begin
      Application.MessageBox('Digite a porta!','Campo porta vazio',0);
      EdtPorta.SetFocus;
      Screen.Cursor:=crDefault;
      Exit;
    end else if trim(EdtAltUser.Text)='' then
    begin
      Application.MessageBox('Digite o nome de usuário para conexão!','Campo usuário vazio',0);
      EdtAltUser.SetFocus;
      Screen.Cursor:=crDefault;
      Exit;
    end;
   try
         Database.HostName:=EdtServ.Text;
         Database.Port:=strtoint(EdtPorta.Text);
         Database.UserName:=EdtAltUser.Text;
         Database.Password:=EdtAltSenha.Text;
         Database.Connected:=true;
         Database.Connected:=false;

         ini:=TINIFile.Create('ControleNFe.ini');
         ini.WriteString('conexao','host',EdtServ.Text);
         ini.WriteInteger('conexao','porta',strtoint(EdtPorta.Text));
         ini.WriteString('conexao','user',EdtAltUser.Text);
         ini.WriteString('conexao','psw',EdtAltSenha.Text);
         ini.Free;
         Application.MessageBox('Servidor de conexão alterado!',PChar('Conectado em: '+EdtServ.Text),0);

   except
         Application.MessageBox(PChar('Não foi possível conectar no servidor "'+EdtServ.Text+'" com os dados indicados.'),'Conexão falhou.',0);
         EdtServ.SetFocus;
         Screen.Cursor:=crDefault;
         Exit;
   end;

  EdtServ.Text:='';
  EdtPorta.Text:='3306';
  EdtAltUser.Text:='';
  EdtAltSenha.Text:='';
  Servidor.Visible:=false;
  Login.Visible:=true;
  EdtUser.SetFocus;
  Screen.Cursor:=crDefault;
end;



procedure TFMenu.BtnVoltarClick(Sender: TObject);
begin
  EdtNovaSenha.Text:='';
  EdtNovoUser.Text:='';
  EdtSenha.Text:='';
  EdtUser.Text:='';
  Cadastro.Visible:=false;
  Login.Visible:=true;
  EdtUser.SetFocus;
end;

procedure TFMenu.EdtNovaSenhaChange(Sender: TObject);
begin

end;


procedure TFMenu.FormKeyPress(Sender: TObject; var Key: char);
begin
  try
     if ActiveControl.ClassName='TButton' then
           Exit;
     if key=#13 then
     begin
       key:=#0;
       SelectNext(ActiveControl,True,True);
     end;
  except
  end;
end;

procedure TFMenu.ItemRelatoriosClick(Sender: TObject);
begin
     Q.Close;
     Q.SQL.Clear;
     Q.SQL.Text:='select not_cha from notas where not_usr='+inttostr(user);
     Q.Open;
     if (Q.RecNo=0) then
     begin
          Application.MessageBox('Não há notas cadastradas','Sem dados para relatório',0);
          Exit;
     end;
     FRel.user:=user;
     FRel.Show;
end;

procedure TFMenu.ItemRemoveClick(Sender: TObject);
begin
  Q.Close;
  Q.SQL.Clear;
  Q.SQL.Text:='select not_cha from notas where not_usr='+inttostr(user);
  Q.Open;
  if (Q.RecNo=0) then
  begin
       Application.MessageBox('Não há notas cadastradas','Sem dados para remover',0);
       Exit;
  end;
  FRemove.user:=user;
  FRemove.Show;
end;

procedure TFMenu.ItemVisualizaClick(Sender: TObject);
begin
  Q.Close;
  Q.SQL.Clear;
  Q.SQL.Text:='select not_cha from notas where not_usr='+inttostr(user);
  Q.Open;
  if (Q.RecNo=0) then
  begin
       Application.MessageBox('Não há notas cadastradas','Sem dados para visualizar',0);
       Exit;
  end;
  FVisualiza.user:=user;
  FVisualiza.Show;
end;

procedure TFMenu.MostraSenhaChange(Sender: TObject);
begin
  if MostraSenha.Checked then
     EdtSenha.PasswordChar:=#0
  else
     EdtSenha.PasswordChar:='*';
end;

procedure TFMenu.MostraSenhaNova1Change(Sender: TObject);
begin
    if MostraSenhaNova1.Checked then
     EdtAltSenha.PasswordChar:=#0
  else
     EdtAltSenha.PasswordChar:='*';
end;

procedure TFMenu.MostraSenhaNovaChange(Sender: TObject);
begin
  if MostraSenhaNova.Checked then
     EdtNovaSenha.PasswordChar:=#0
  else
     EdtNovaSenha.PasswordChar:='*';
end;

procedure TFMenu.FormActivate(Sender: TObject);
begin
  if Login.Visible then
     EdtUser.SetFocus
  else if Cadastro.Visible then
     EdtNovoUser.SetFocus;
end;

procedure TFMenu.FormCreate(Sender: TObject);
begin
  DefaultFormatSettings.LongMonthNames[3]:='março';
  DefaultFormatSettings.LongDayNames[1]:='domingo';
  DefaultFormatSettings.LongdayNames[2]:='segunda-feira';
  DefaultFormatSettings.LongDayNames[3]:='terça-feira';
  DefaultFormatSettings.LongDayNames[4]:='quarta-feira';
  DefaultFormatSettings.LongDayNames[5]:='quinta-feira';
  DefaultFormatSettings.LongdayNames[6]:='sexta-feira';
  DefaultFormatSettings.LongDayNames[7]:='sábado';

  MenuItens.Items[0].Visible:=false;
  MenuItens.Items[1].Visible:=false;
  MenuItens.Items[2].Visible:=false;
  MenuItens.Items[3].Visible:=false;
  MenuItens.Items[4].Visible:=false;
  MenuItens.Items[5].Visible:=false;
  StatusBar.Visible:=false;

  Login.Top:=80;
  Login.Left:=136;
  Cadastro.Top:=80;
  Cadastro.Left:=136;
  Servidor.Top:=80;
  Servidor.Left:=136;
  FMenu.Height:=455;
  FMenu.Width:=735;


  Login.Visible:=true;
  Botoes.Visible:=false;
end;



end.

