VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} de_informaEM 
   ClientHeight    =   7260
   ClientLeft      =   7515
   ClientTop       =   1455
   ClientWidth     =   3570
   _ExtentX        =   6297
   _ExtentY        =   12806
   FolderFlags     =   5
   TypeLibGuid     =   "{237E11FB-5E4C-4F00-AC9B-66666E2D8B0C}"
   TypeInfoGuid    =   "{7383E427-E9F8-4C88-8DCD-3305AA122CB1}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "cn_informa"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=SQLOLEDB.1;Password=zx11bbb7;Persist Security Info=True;User ID=sa;Initial Catalog=LRF;Data Source=(local)"
      IsSQL           =   -1  'True
      QuoteChar       =   34
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   172
   BeginProperty Recordset1 
      CommandName     =   "Sel_Projetos"
      CommDispId      =   1002
      RsDispId        =   1007
      CommandText     =   "select * from tb_cadprojetos where status like ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "projeto"
         Caption         =   "projeto"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field4 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "StatusLike"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "Ins_CadProjetos"
      CommDispId      =   1008
      RsDispId        =   -1
      CommandText     =   "insert into tb_cadprojetos (projeto, descricao, status, datacad, usuariocad) values ( ? , ? , ? , ? , ? ) "
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   5
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "Descricao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   80
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "Status"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "DataCad"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "UsuarioCad"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "Alt_CadProjetos"
      CommDispId      =   1018
      RsDispId        =   -1
      CommandText     =   "update tb_cadprojetos set descricao = ?, status = ? where projeto = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Descricao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   80
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "Status"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "Projeto"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "Sel_DataServidor"
      CommDispId      =   1020
      RsDispId        =   1025
      CommandText     =   "select getdate() agora"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "agora"
         Caption         =   "agora"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset5 
      CommandName     =   "Sel_ProjetosNome"
      CommDispId      =   1026
      RsDispId        =   1032
      CommandText     =   "select * from tb_cadprojetos where projeto = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "projeto"
         Caption         =   "projeto"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field4 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Projeto"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset6 
      CommandName     =   "Sel_CadSubContCGC"
      CommDispId      =   1033
      RsDispId        =   1038
      CommandText     =   "select * from tb_cadsubcontra where cgc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   16
      BeginProperty Field1 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "razaosoc"
         Caption         =   "razaosoc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "fantasia"
         Caption         =   "fantasia"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "apelido"
         Caption         =   "apelido"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "cep"
         Caption         =   "cep"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "endereco"
         Caption         =   "endereco"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "ie"
         Caption         =   "ie"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "fones"
         Caption         =   "fones"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "fax"
         Caption         =   "fax"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "contatos"
         Caption         =   "contatos"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "tiposub"
         Caption         =   "tiposub"
      EndProperty
      BeginProperty Field15 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "CGC"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset7 
      CommandName     =   "Ins_CadSubContra"
      CommDispId      =   1039
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":0000
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   17
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   50
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "Param11"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "Param12"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "Param13"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   50
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "Param14"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "Param15"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "Param16"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "Param17"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset8 
      CommandName     =   "Alt_CadSubContra"
      CommDispId      =   1066
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":00F2
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   14
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   50
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "Param11"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "Param12"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   50
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "Param13"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "Param14"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset9 
      CommandName     =   "Sel_CidadeCEP"
      CommDispId      =   1068
      RsDispId        =   2050
      CommandText     =   "select * from tb_cadcidades where cepvali >= ? and cepvalf <= ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "cepi"
         Caption         =   "cepi"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "cepf"
         Caption         =   "cepf"
      EndProperty
      BeginProperty Field3 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "cepvali"
         Caption         =   "cepvali"
      EndProperty
      BeginProperty Field4 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "cepvalf"
         Caption         =   "cepvalf"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "cim"
         Caption         =   "cim"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset10 
      CommandName     =   "Alt_BloqSubContra"
      CommDispId      =   1074
      RsDispId        =   -1
      CommandText     =   "update tb_cadsubcontra set status = 0 where cgc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "CGC"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset11 
      CommandName     =   "Sel_CadCliCGC"
      CommDispId      =   1076
      RsDispId        =   1572
      CommandText     =   "select * from tb_cadcli where cgc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   53
      BeginProperty Field1 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "nome"
         Caption         =   "nome"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fantasia"
         Caption         =   "fantasia"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "apelido"
         Caption         =   "apelido"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "endereco"
         Caption         =   "endereco"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "cep"
         Caption         =   "cep"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "ie"
         Caption         =   "ie"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "pabx"
         Caption         =   "pabx"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fax"
         Caption         =   "fax"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "contato1"
         Caption         =   "contato1"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fonecontato1"
         Caption         =   "fonecontato1"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "emailcontato1"
         Caption         =   "emailcontato1"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "anivercontato1"
         Caption         =   "anivercontato1"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "avisarcontato1"
         Caption         =   "avisarcontato1"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "avusucontato1"
         Caption         =   "avusucontato1"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "contato2"
         Caption         =   "contato2"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fonecontato2"
         Caption         =   "fonecontato2"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "emailcontato2"
         Caption         =   "emailcontato2"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "anivercontato2"
         Caption         =   "anivercontato2"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "avisarcontato2"
         Caption         =   "avisarcontato2"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "avusucontato2"
         Caption         =   "avusucontato2"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "contato3"
         Caption         =   "contato3"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fonecontato3"
         Caption         =   "fonecontato3"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "emailcontato3"
         Caption         =   "emailcontato3"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "anivercontato3"
         Caption         =   "anivercontato3"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "avisarcontato3"
         Caption         =   "avisarcontato3"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "avusucontato3"
         Caption         =   "avusucontato3"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "consigentrega"
         Caption         =   "consigentrega"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "consigentregaair"
         Caption         =   "consigentregaair"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "consigtransf"
         Caption         =   "consigtransf"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "consigdevol"
         Caption         =   "consigdevol"
      EndProperty
      BeginProperty Field35 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "atendusu"
         Caption         =   "atendusu"
      EndProperty
      BeginProperty Field36 
         Precision       =   0
         Size            =   6
         Scale           =   0
         Type            =   200
         Name            =   "prazo"
         Caption         =   "prazo"
      EndProperty
      BeginProperty Field37 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      BeginProperty Field38 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field39 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ultemissao"
         Caption         =   "ultemissao"
      EndProperty
      BeginProperty Field40 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field41 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "databloqueio"
         Caption         =   "databloqueio"
      EndProperty
      BeginProperty Field42 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usubloqueio"
         Caption         =   "usubloqueio"
      EndProperty
      BeginProperty Field43 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "descrbloqueio"
         Caption         =   "descrbloqueio"
      EndProperty
      BeginProperty Field44 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "env_emailOco"
         Caption         =   "env_emailOco"
      EndProperty
      BeginProperty Field45 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "env_emailFer"
         Caption         =   "env_emailFer"
      EndProperty
      BeginProperty Field46 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email1"
         Caption         =   "email1"
      EndProperty
      BeginProperty Field47 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email2"
         Caption         =   "email2"
      EndProperty
      BeginProperty Field48 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email3"
         Caption         =   "email3"
      EndProperty
      BeginProperty Field49 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email4"
         Caption         =   "email4"
      EndProperty
      BeginProperty Field50 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email5"
         Caption         =   "email5"
      EndProperty
      BeginProperty Field51 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "rem_des_log"
         Caption         =   "rem_des_log"
      EndProperty
      BeginProperty Field52 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "alarm_ger"
         Caption         =   "alarm_ger"
      EndProperty
      BeginProperty Field53 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "hosp"
         Caption         =   "hosp"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset12 
      CommandName     =   "Sel_CadCidadePorCidadeUF"
      CommDispId      =   1084
      RsDispId        =   1823
      CommandText     =   "select * from tb_cadcidades where cidade = ? and uf = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "cepi"
         Caption         =   "cepi"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "cepf"
         Caption         =   "cepf"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "cim"
         Caption         =   "cim"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "UF"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset13 
      CommandName     =   "Sel_CidadesLike"
      CommDispId      =   1094
      RsDispId        =   1824
      CommandText     =   "select cidade, uf from tb_cadcidades where cidade like ? order by cidade"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset14 
      CommandName     =   "Ins_CadCli"
      CommDispId      =   1105
      RsDispId        =   -1
      CommandText     =   "dbo.sp_ins_cadcli"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CallSyntax      =   $"de_informaEM.dsx":01B5
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   40
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@fantasia"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@apelido"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@endereco"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@complemento"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@cep"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@ie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "@pabx"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "@fax"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "@contato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "@fonecontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "@emailcontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "@anivercontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "@avisarcontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "@avusucontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "@contato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "@fonecontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P22 
         RealName        =   "@emailcontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P23 
         RealName        =   "@anivercontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P24 
         RealName        =   "@avisarcontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P25 
         RealName        =   "@avusucontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P26 
         RealName        =   "@contato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P27 
         RealName        =   "@fonecontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P28 
         RealName        =   "@emailcontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P29 
         RealName        =   "@anivercontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P30 
         RealName        =   "@avisarcontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P31 
         RealName        =   "@avusucontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P32 
         RealName        =   "@consigentrega"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P33 
         RealName        =   "@consigentregaair"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P34 
         RealName        =   "@consigtransf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P35 
         RealName        =   "@consigdevol"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P36 
         RealName        =   "@atendusu"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P37 
         RealName        =   "@prazo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   6
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P38 
         RealName        =   "@usuariocad"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P39 
         RealName        =   "@rem_des_log"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P40 
         RealName        =   "@alarm_ger"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset15 
      CommandName     =   "Sel_CadUsuarioPorUsu"
      CommDispId      =   1107
      RsDispId        =   1112
      CommandText     =   "select usuario, nome, status from tb_cadusu where usuario = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuario"
         Caption         =   "usuario"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "nome"
         Caption         =   "nome"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Usuario"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset16 
      CommandName     =   "Sel_CadCliCGCLike"
      CommDispId      =   1113
      RsDispId        =   1616
      CommandText     =   "select cgc, nome, fantasia, apelido, cidade, uf, rem_des_log, endereco, complemento from tb_cadcli where cgc like ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   9
      BeginProperty Field1 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "nome"
         Caption         =   "nome"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fantasia"
         Caption         =   "fantasia"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "apelido"
         Caption         =   "apelido"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "rem_des_log"
         Caption         =   "rem_des_log"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "endereco"
         Caption         =   "endereco"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset17 
      CommandName     =   "Sel_CadCliRazaoLike"
      CommDispId      =   1123
      RsDispId        =   1617
      CommandText     =   "select cgc, nome, fantasia, apelido, cidade, uf, rem_des_log, endereco, complemento from tb_cadcli where nome like ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   9
      BeginProperty Field1 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "nome"
         Caption         =   "nome"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fantasia"
         Caption         =   "fantasia"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "apelido"
         Caption         =   "apelido"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "rem_des_log"
         Caption         =   "rem_des_log"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "endereco"
         Caption         =   "endereco"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset18 
      CommandName     =   "Sel_CadCliFantasiaLike"
      CommDispId      =   1130
      RsDispId        =   1618
      CommandText     =   "select cgc, nome, fantasia, apelido, cidade, uf, rem_des_log, endereco, complemento from tb_cadcli where fantasia like ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   9
      BeginProperty Field1 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "nome"
         Caption         =   "nome"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fantasia"
         Caption         =   "fantasia"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "apelido"
         Caption         =   "apelido"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "rem_des_log"
         Caption         =   "rem_des_log"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "endereco"
         Caption         =   "endereco"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "fantasia"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset19 
      CommandName     =   "Alt_CadClientes"
      CommDispId      =   1136
      RsDispId        =   -1
      CommandText     =   "dbo.sp_alt_cadclientes"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CallSyntax      =   $"de_informaEM.dsx":024C
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   39
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@fantasia"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@apelido"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@endereco"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@complemento"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@cep"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@ie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "@pabx"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "@fax"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "@contato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "@fonecontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "@emailcontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "@anivercontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "@avisarcontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "@avusucontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "@contato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "@fonecontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P22 
         RealName        =   "@emailcontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P23 
         RealName        =   "@anivercontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P24 
         RealName        =   "@avisarcontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P25 
         RealName        =   "@avusucontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P26 
         RealName        =   "@contato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P27 
         RealName        =   "@fonecontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P28 
         RealName        =   "@emailcontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P29 
         RealName        =   "@anivercontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P30 
         RealName        =   "@avisarcontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P31 
         RealName        =   "@avusucontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P32 
         RealName        =   "@consigentrega"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P33 
         RealName        =   "@consigentregaair"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P34 
         RealName        =   "@consigtransf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P35 
         RealName        =   "@consigdevol"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P36 
         RealName        =   "@atendusu"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P37 
         RealName        =   "@prazo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   6
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P38 
         RealName        =   "@rem_des_log"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P39 
         RealName        =   "@alarm_ger"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset20 
      CommandName     =   "Alt_BloqueioCliente"
      CommDispId      =   1138
      RsDispId        =   -1
      CommandText     =   "update tb_cadcli set status = ?, databloqueio = getdate(), usubloqueio = ?, descrbloqueio = ?  where cgc LIKE ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   4
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "STATUS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "usuario"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "descr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   50
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "cgclike"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset21 
      CommandName     =   "Ins_CadCliHistorico"
      CommDispId      =   1144
      RsDispId        =   -1
      CommandText     =   "insert Into tb_cadclihist (cgc, data, descricao, usuario) values ( ? , getdate() , ? , ? )"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "CGC"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "Descr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   100
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "Usuario"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset22 
      CommandName     =   "Sel_CadCliHistorico"
      CommDispId      =   1146
      RsDispId        =   1152
      CommandText     =   "select * from tb_cadclihist where cgc = ? order by data desc"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field2 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuario"
         Caption         =   "usuario"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "CGC"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset23 
      CommandName     =   "Sel_CadCliProds"
      CommDispId      =   1153
      RsDispId        =   1540
      CommandText     =   "select * from tb_cadcliprods where cgc = ? and status = '1' and natproduto like ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "natproduto"
         Caption         =   "natproduto"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "classiata"
         Caption         =   "classiata"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "obspadrao"
         Caption         =   "obspadrao"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usucad"
         Caption         =   "usucad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset24 
      CommandName     =   "Ins_CadCliProds"
      CommDispId      =   1159
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":02E5
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   5
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "CGC"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "Prod"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "Iata"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "ObsPadrao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   50
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "Usuario"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset25 
      CommandName     =   "Sel_ClassIATAPorCod"
      CommDispId      =   1168
      RsDispId        =   1173
      CommandText     =   "select * from tb_aircadiata where codigo = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Cod"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset26 
      CommandName     =   "Alt_CadCliProds"
      CommDispId      =   1174
      RsDispId        =   -1
      CommandText     =   "update tb_cadcliprods set classiata = ?, obspadrao = ? where cgc = ? and natproduto = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   4
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Iata"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "Obs"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   50
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "CGC"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "Produto"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset27 
      CommandName     =   "Sel_CadUfs"
      CommDispId      =   1179
      RsDispId        =   1186
      CommandText     =   "select * from tb_cadufs where uf = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "nome"
         Caption         =   "nome"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset28 
      CommandName     =   "Sel_TempTR01"
      CommDispId      =   1187
      RsDispId        =   1246
      CommandText     =   "select * from tb_temptr01 where codigo = ? order by data"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "cim"
         Caption         =   "cim"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretemin"
         Caption         =   "fretemin"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "tarifaperc"
         Caption         =   "tarifaperc"
      EndProperty
      BeginProperty Field10 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset29 
      CommandName     =   "Ins_TempTR01"
      CommDispId      =   1193
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":0372
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   9
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "descricao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   80
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "origem"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "origemuf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "cim"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         UserName        =   "cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         UserName        =   "fretemin"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         UserName        =   "tarifaperc"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset30 
      CommandName     =   "Sel_CadUfsTodos"
      CommDispId      =   1204
      RsDispId        =   1210
      CommandText     =   "select * from tb_cadufs order by regiaogeo"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "nome"
         Caption         =   "nome"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset31 
      CommandName     =   "Sel_UFnaoTratTR01"
      CommDispId      =   1218
      RsDispId        =   1229
      CommandText     =   "dbo.sp_sel_ufsnaotratTR01"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CallSyntax      =   "{? = CALL dbo.sp_sel_ufsnaotratTR01( ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "nome"
         Caption         =   "nome"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset32 
      CommandName     =   "Sel_TempTR01Confere"
      CommDispId      =   1230
      RsDispId        =   1239
      CommandText     =   "select * from tb_temptr01 where codigo = ? and uf = ? and cim = ? and cidade = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   4
         Scale           =   0
         Type            =   200
         Name            =   "cim"
         Caption         =   "cim"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretemin"
         Caption         =   "fretemin"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "tarifaperc"
         Caption         =   "tarifaperc"
      EndProperty
      NumGroups       =   0
      ParamCount      =   4
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "UF"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "CIM"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "Cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset33 
      CommandName     =   "Exc_TempTR01localidade"
      CommDispId      =   1240
      RsDispId        =   -1
      CommandText     =   "delete from tb_temptr01 where codigo = ? and uf = ? and cim = ? and cidade = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   4
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "UF"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "CIM"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "Cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset34 
      CommandName     =   "Exc_TempTR01Tudo"
      CommDispId      =   1242
      RsDispId        =   -1
      CommandText     =   "delete from tb_temptr01 where codigo = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset35 
      CommandName     =   "Ins_TR01Oficial"
      CommDispId      =   1244
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":0418
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   12
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "descricao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   80
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "origem"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "origemuf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "cim"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         UserName        =   "cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         UserName        =   "fretemin"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         UserName        =   "tarifaperc"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         UserName        =   "inicvigencia"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "Param11"
         UserName        =   "usuario"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "Param12"
         UserName        =   "datacad"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset36 
      CommandName     =   "Sel_TR01"
      CommDispId      =   1251
      RsDispId        =   1361
      CommandText     =   $"de_informaEM.dsx":04FB
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field8 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field10 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset37 
      CommandName     =   "Sel_TR01Expiradas"
      CommDispId      =   1257
      RsDispId        =   1268
      CommandText     =   $"de_informaEM.dsx":05B1
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field5 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field8 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset38 
      CommandName     =   "Sel_TR01Codigo"
      CommDispId      =   1269
      RsDispId        =   1275
      CommandText     =   "select uf, cim, cidade, fretemin, tarifaperc from tb_tr01 where codigo = ? order by uf, cim"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   4
         Scale           =   0
         Type            =   200
         Name            =   "cim"
         Caption         =   "cim"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretemin"
         Caption         =   "fretemin"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "tarifaperc"
         Caption         =   "tarifaperc"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset39 
      CommandName     =   "Sel_CadLocalidadeSigla"
      CommDispId      =   1278
      RsDispId        =   1286
      CommandText     =   "select * from tb_aircadlocal where sigla like ? order by localidade"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "sigla"
         Caption         =   "sigla"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "localidade"
         Caption         =   "localidade"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset40 
      CommandName     =   "Ins_TempTA01"
      CommDispId      =   1280
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":064B
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   20
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "descricao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   80
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "origemuf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "origem"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "localidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "sigla"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         UserName        =   "uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         UserName        =   "txminima"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         UserName        =   "porkilo"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         UserName        =   "advalor"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "Param11"
         UserName        =   "coletaate"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "Param12"
         UserName        =   "coletavalor"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "Param13"
         UserName        =   "coletaexced"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "Param14"
         UserName        =   "entregaate"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "Param15"
         UserName        =   "entregavalor"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "Param16"
         UserName        =   "entregaexced"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "Param17"
         UserName        =   "regiaogeo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "Param18"
         UserName        =   "redespate"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "Param19"
         UserName        =   "redespvalor"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "Param20"
         UserName        =   "redespexced"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset41 
      CommandName     =   "Sel_TempTA01Cod"
      CommDispId      =   1295
      RsDispId        =   1374
      CommandText     =   $"de_informaEM.dsx":07C8
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   20
      BeginProperty Field1 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "sigla"
         Caption         =   "sigla"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "localidade"
         Caption         =   "localidade"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txminima"
         Caption         =   "txminima"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "porkilo"
         Caption         =   "porkilo"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_advalorem"
         Caption         =   "gen_advalorem"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaate"
         Caption         =   "gen_txcoletaate"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletavalor"
         Caption         =   "gen_txcoletavalor"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaexced"
         Caption         =   "gen_txcoletaexced"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaate"
         Caption         =   "gen_txentregaate"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregavalor"
         Caption         =   "gen_txentregavalor"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaexced"
         Caption         =   "gen_txentregaexced"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      BeginProperty Field18 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_ate"
         Caption         =   "txredesp_ate"
      EndProperty
      BeginProperty Field19 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_valor"
         Caption         =   "txredesp_valor"
      EndProperty
      BeginProperty Field20 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_exced"
         Caption         =   "txredesp_exced"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset42 
      CommandName     =   "Sel_TempTA01Confere"
      CommDispId      =   1301
      RsDispId        =   1306
      CommandText     =   "select * from tb_tempta01 where codigo = ? and sigla = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   18
      BeginProperty Field1 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "localidade"
         Caption         =   "localidade"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "sigla"
         Caption         =   "sigla"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txminima"
         Caption         =   "txminima"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "porkilo"
         Caption         =   "porkilo"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_advalorem"
         Caption         =   "gen_advalorem"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaate"
         Caption         =   "gen_txcoletaate"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletavalor"
         Caption         =   "gen_txcoletavalor"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaexced"
         Caption         =   "gen_txcoletaexced"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaate"
         Caption         =   "gen_txentregaate"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregavalor"
         Caption         =   "gen_txentregavalor"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaexced"
         Caption         =   "gen_txentregaexced"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_ate"
         Caption         =   "txredesp_ate"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_valor"
         Caption         =   "txredesp_valor"
      EndProperty
      BeginProperty Field18 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_exced"
         Caption         =   "txredesp_exced"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "sigla"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset43 
      CommandName     =   "Exc_TempTA01Localidade"
      CommDispId      =   1307
      RsDispId        =   -1
      CommandText     =   "delete from tb_tempta01 where codigo = ? and sigla = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "Sigla"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset44 
      CommandName     =   "Exc_TempTA01Tudo"
      CommDispId      =   1315
      RsDispId        =   -1
      CommandText     =   "delete from tb_tempta01 where codigo = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset45 
      CommandName     =   "Ins_TA01Oficial"
      CommDispId      =   1317
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":08FC
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   23
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "cod"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "descr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   80
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "origemuf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "origem"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "local"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "sigla"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         UserName        =   "uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         UserName        =   "txmin"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         UserName        =   "kilo"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         UserName        =   "advalor"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "Param11"
         UserName        =   "coletaate"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "Param12"
         UserName        =   "coletavalor"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "Param13"
         UserName        =   "coletaexced"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "Param14"
         UserName        =   "entregaate"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "Param15"
         UserName        =   "entregavalor"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "Param16"
         UserName        =   "entregaexced"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "Param17"
         UserName        =   "regiaogeo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "Param18"
         UserName        =   "redespate"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "Param19"
         UserName        =   "redespvalor"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "Param20"
         UserName        =   "redespexced"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "Param21"
         UserName        =   "inicvigencia"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P22 
         RealName        =   "Param22"
         UserName        =   "usuario"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P23 
         RealName        =   "Param23"
         UserName        =   "datacad"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset46 
      CommandName     =   "Sel_TA01"
      CommDispId      =   1319
      RsDispId        =   1368
      CommandText     =   $"de_informaEM.dsx":0AC9
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field8 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field10 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset47 
      CommandName     =   "Sel_TA01Expiradas"
      CommDispId      =   1325
      RsDispId        =   1330
      CommandText     =   $"de_informaEM.dsx":0B7E
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field5 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field8 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset48 
      CommandName     =   "Sel_TA01Codigo"
      CommDispId      =   1340
      RsDispId        =   1345
      CommandText     =   $"de_informaEM.dsx":0C18
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   20
      BeginProperty Field1 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "sigla"
         Caption         =   "sigla"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "localidade"
         Caption         =   "localidade"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txminima"
         Caption         =   "txminima"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "porkilo"
         Caption         =   "porkilo"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_advalorem"
         Caption         =   "gen_advalorem"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaate"
         Caption         =   "gen_txcoletaate"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletavalor"
         Caption         =   "gen_txcoletavalor"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaexced"
         Caption         =   "gen_txcoletaexced"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaate"
         Caption         =   "gen_txentregaate"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregavalor"
         Caption         =   "gen_txentregavalor"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaexced"
         Caption         =   "gen_txentregaexced"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_ate"
         Caption         =   "txredesp_ate"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_valor"
         Caption         =   "txredesp_valor"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_exced"
         Caption         =   "txredesp_exced"
      EndProperty
      BeginProperty Field17 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field18 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field20 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset49 
      CommandName     =   "Ins_TG01"
      CommDispId      =   1392
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":0D69
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   10
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "descricao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   80
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "origemuf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "origem"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "fretepeso"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "freteminimo"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         UserName        =   "fretevalor"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         UserName        =   "inicvigencia"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         UserName        =   "usuariocad"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         UserName        =   "datacad"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset50 
      CommandName     =   "Sel_TG01"
      CommDispId      =   1394
      RsDispId        =   1399
      CommandText     =   $"de_informaEM.dsx":0E43
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field8 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field10 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset51 
      CommandName     =   "Sel_TG01Codigo"
      CommDispId      =   1400
      RsDispId        =   1405
      CommandText     =   "select fretepeso, freteminimo, fretevalor from tb_tg01 where codigo = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepeso"
         Caption         =   "fretepeso"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "freteminimo"
         Caption         =   "freteminimo"
      EndProperty
      BeginProperty Field3 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalor"
         Caption         =   "fretevalor"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset52 
      CommandName     =   "Sel_TempTA02Cod"
      CommDispId      =   1406
      RsDispId        =   1806
      CommandText     =   $"de_informaEM.dsx":0EF8
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   16
      BeginProperty Field1 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesode"
         Caption         =   "pesode"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesoate"
         Caption         =   "pesoate"
      EndProperty
      BeginProperty Field3 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valormin"
         Caption         =   "valormin"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "porkilo"
         Caption         =   "porkilo"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_advalorem"
         Caption         =   "gen_advalorem"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaate"
         Caption         =   "gen_txcoletaate"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletavalor"
         Caption         =   "gen_txcoletavalor"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaexced"
         Caption         =   "gen_txcoletaexced"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaate"
         Caption         =   "gen_txentregaate"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregavalor"
         Caption         =   "gen_txentregavalor"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaexced"
         Caption         =   "gen_txentregaexced"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset53 
      CommandName     =   "Sel_TempTA02Confere"
      CommDispId      =   1412
      RsDispId        =   1794
      CommandText     =   "select * from tb_tempta02 where codigo = ? and pesode = ? and pesoate = ? "
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   16
      BeginProperty Field1 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesode"
         Caption         =   "pesode"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesoate"
         Caption         =   "pesoate"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valormin"
         Caption         =   "valormin"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "porkilo"
         Caption         =   "porkilo"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_advalorem"
         Caption         =   "gen_advalorem"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaate"
         Caption         =   "gen_txcoletaate"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletavalor"
         Caption         =   "gen_txcoletavalor"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaexced"
         Caption         =   "gen_txcoletaexced"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaate"
         Caption         =   "gen_txentregaate"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregavalor"
         Caption         =   "gen_txentregavalor"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaexced"
         Caption         =   "gen_txentregaexced"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "Pesode"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "PesoAte"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset54 
      CommandName     =   "Exc_TempTA02Linha"
      CommDispId      =   1418
      RsDispId        =   -1
      CommandText     =   "delete from tb_tempta02 where codigo = ? and pesode = ? and pesoate = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "Pesode"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "Pesoate"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset55 
      CommandName     =   "Exc_TempTA02Tudo"
      CommDispId      =   1427
      RsDispId        =   -1
      CommandText     =   "delete from tb_tempta02 where codigo = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset56 
      CommandName     =   "Ins_TA02Oficial"
      CommDispId      =   1433
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":0FFB
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   19
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   100
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "Param11"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "Param12"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "Param13"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "Param14"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "Param15"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "Param16"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "Param17"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "Param18"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "Param19"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset57 
      CommandName     =   "Sel_TA02"
      CommDispId      =   1450
      RsDispId        =   1751
      CommandText     =   $"de_informaEM.dsx":1186
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field8 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field10 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset58 
      CommandName     =   "Sel_TA02Codigo"
      CommDispId      =   1456
      RsDispId        =   1801
      CommandText     =   $"de_informaEM.dsx":123A
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   16
      BeginProperty Field1 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesode"
         Caption         =   "pesode"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesoate"
         Caption         =   "pesoate"
      EndProperty
      BeginProperty Field3 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valormin"
         Caption         =   "valormin"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "porkilo"
         Caption         =   "porkilo"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_advalorem"
         Caption         =   "gen_advalorem"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaate"
         Caption         =   "gen_txcoletaate"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletavalor"
         Caption         =   "gen_txcoletavalor"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaexced"
         Caption         =   "gen_txcoletaexced"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaate"
         Caption         =   "gen_txentregaate"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregavalor"
         Caption         =   "gen_txentregavalor"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaexced"
         Caption         =   "gen_txentregaexced"
      EndProperty
      BeginProperty Field13 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field14 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field16 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset59 
      CommandName     =   "Ins_TempTA02"
      CommDispId      =   1462
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":1366
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   16
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   100
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "Param11"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "Param12"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "Param13"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "Param14"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "Param15"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "Param16"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset60 
      CommandName     =   "Sel_CadCliFantasia"
      CommDispId      =   1466
      RsDispId        =   1472
      CommandText     =   "select * from tb_cadcli where fantasia = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   52
      BeginProperty Field1 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "nome"
         Caption         =   "nome"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fantasia"
         Caption         =   "fantasia"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "apelido"
         Caption         =   "apelido"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "endereco"
         Caption         =   "endereco"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "cep"
         Caption         =   "cep"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "ie"
         Caption         =   "ie"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "pabx"
         Caption         =   "pabx"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fax"
         Caption         =   "fax"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "contato1"
         Caption         =   "contato1"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fonecontato1"
         Caption         =   "fonecontato1"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "emailcontato1"
         Caption         =   "emailcontato1"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "anivercontato1"
         Caption         =   "anivercontato1"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "avisarcontato1"
         Caption         =   "avisarcontato1"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "avusucontato1"
         Caption         =   "avusucontato1"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "contato2"
         Caption         =   "contato2"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fonecontato2"
         Caption         =   "fonecontato2"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "emailcontato2"
         Caption         =   "emailcontato2"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "anivercontato2"
         Caption         =   "anivercontato2"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "avisarcontato2"
         Caption         =   "avisarcontato2"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "avusucontato2"
         Caption         =   "avusucontato2"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "contato3"
         Caption         =   "contato3"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fonecontato3"
         Caption         =   "fonecontato3"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "emailcontato3"
         Caption         =   "emailcontato3"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "anivercontato3"
         Caption         =   "anivercontato3"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "avisarcontato3"
         Caption         =   "avisarcontato3"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "avusucontato3"
         Caption         =   "avusucontato3"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "consigentrega"
         Caption         =   "consigentrega"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "consigtransf"
         Caption         =   "consigtransf"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "consigdevol"
         Caption         =   "consigdevol"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "atendusu"
         Caption         =   "atendusu"
      EndProperty
      BeginProperty Field35 
         Precision       =   0
         Size            =   6
         Scale           =   0
         Type            =   200
         Name            =   "prazo"
         Caption         =   "prazo"
      EndProperty
      BeginProperty Field36 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      BeginProperty Field37 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field38 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ultemissao"
         Caption         =   "ultemissao"
      EndProperty
      BeginProperty Field39 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field40 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "databloqueio"
         Caption         =   "databloqueio"
      EndProperty
      BeginProperty Field41 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usubloqueio"
         Caption         =   "usubloqueio"
      EndProperty
      BeginProperty Field42 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "descrbloqueio"
         Caption         =   "descrbloqueio"
      EndProperty
      BeginProperty Field43 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "env_emailOco"
         Caption         =   "env_emailOco"
      EndProperty
      BeginProperty Field44 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "env_emailFer"
         Caption         =   "env_emailFer"
      EndProperty
      BeginProperty Field45 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email1"
         Caption         =   "email1"
      EndProperty
      BeginProperty Field46 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email2"
         Caption         =   "email2"
      EndProperty
      BeginProperty Field47 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email3"
         Caption         =   "email3"
      EndProperty
      BeginProperty Field48 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email4"
         Caption         =   "email4"
      EndProperty
      BeginProperty Field49 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email5"
         Caption         =   "email5"
      EndProperty
      BeginProperty Field50 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "rem_des_log"
         Caption         =   "rem_des_log"
      EndProperty
      BeginProperty Field51 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "alarm_ger"
         Caption         =   "alarm_ger"
      EndProperty
      BeginProperty Field52 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "hosp"
         Caption         =   "hosp"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Fantasia"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset61 
      CommandName     =   "Sel_CadCliApelido"
      CommDispId      =   1473
      RsDispId        =   1479
      CommandText     =   "select * from tb_cadcli where apelido = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   52
      BeginProperty Field1 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "nome"
         Caption         =   "nome"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fantasia"
         Caption         =   "fantasia"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "apelido"
         Caption         =   "apelido"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "endereco"
         Caption         =   "endereco"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "cep"
         Caption         =   "cep"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "ie"
         Caption         =   "ie"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "pabx"
         Caption         =   "pabx"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fax"
         Caption         =   "fax"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "contato1"
         Caption         =   "contato1"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fonecontato1"
         Caption         =   "fonecontato1"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "emailcontato1"
         Caption         =   "emailcontato1"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "anivercontato1"
         Caption         =   "anivercontato1"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "avisarcontato1"
         Caption         =   "avisarcontato1"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "avusucontato1"
         Caption         =   "avusucontato1"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "contato2"
         Caption         =   "contato2"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fonecontato2"
         Caption         =   "fonecontato2"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "emailcontato2"
         Caption         =   "emailcontato2"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "anivercontato2"
         Caption         =   "anivercontato2"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "avisarcontato2"
         Caption         =   "avisarcontato2"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "avusucontato2"
         Caption         =   "avusucontato2"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "contato3"
         Caption         =   "contato3"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "fonecontato3"
         Caption         =   "fonecontato3"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "emailcontato3"
         Caption         =   "emailcontato3"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "anivercontato3"
         Caption         =   "anivercontato3"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "avisarcontato3"
         Caption         =   "avisarcontato3"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "avusucontato3"
         Caption         =   "avusucontato3"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "consigentrega"
         Caption         =   "consigentrega"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "consigtransf"
         Caption         =   "consigtransf"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "consigdevol"
         Caption         =   "consigdevol"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "atendusu"
         Caption         =   "atendusu"
      EndProperty
      BeginProperty Field35 
         Precision       =   0
         Size            =   6
         Scale           =   0
         Type            =   200
         Name            =   "prazo"
         Caption         =   "prazo"
      EndProperty
      BeginProperty Field36 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      BeginProperty Field37 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field38 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ultemissao"
         Caption         =   "ultemissao"
      EndProperty
      BeginProperty Field39 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field40 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "databloqueio"
         Caption         =   "databloqueio"
      EndProperty
      BeginProperty Field41 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usubloqueio"
         Caption         =   "usubloqueio"
      EndProperty
      BeginProperty Field42 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "descrbloqueio"
         Caption         =   "descrbloqueio"
      EndProperty
      BeginProperty Field43 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "env_emailOco"
         Caption         =   "env_emailOco"
      EndProperty
      BeginProperty Field44 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "env_emailFer"
         Caption         =   "env_emailFer"
      EndProperty
      BeginProperty Field45 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email1"
         Caption         =   "email1"
      EndProperty
      BeginProperty Field46 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email2"
         Caption         =   "email2"
      EndProperty
      BeginProperty Field47 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email3"
         Caption         =   "email3"
      EndProperty
      BeginProperty Field48 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email4"
         Caption         =   "email4"
      EndProperty
      BeginProperty Field49 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "email5"
         Caption         =   "email5"
      EndProperty
      BeginProperty Field50 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "rem_des_log"
         Caption         =   "rem_des_log"
      EndProperty
      BeginProperty Field51 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "alarm_ger"
         Caption         =   "alarm_ger"
      EndProperty
      BeginProperty Field52 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "hosp"
         Caption         =   "hosp"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Apelido"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset62 
      CommandName     =   "Sel_CadSubContraFantasia"
      CommDispId      =   1480
      RsDispId        =   1486
      CommandText     =   "select * from tb_cadsubcontra where fantasia = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   17
      BeginProperty Field1 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "razaosoc"
         Caption         =   "razaosoc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "fantasia"
         Caption         =   "fantasia"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "apelido"
         Caption         =   "apelido"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "cep"
         Caption         =   "cep"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "endereco"
         Caption         =   "endereco"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "ie"
         Caption         =   "ie"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "fones"
         Caption         =   "fones"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "fax"
         Caption         =   "fax"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "contatos"
         Caption         =   "contatos"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "tiposub"
         Caption         =   "tiposub"
      EndProperty
      BeginProperty Field15 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "fantasia"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset63 
      CommandName     =   "Sel_CadSubContraApelido"
      CommDispId      =   1481
      RsDispId        =   2104
      CommandText     =   "select * from tb_cadsubcontra where apelido = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   17
      BeginProperty Field1 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "nome"
         Caption         =   "nome"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "fantasia"
         Caption         =   "fantasia"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "apelido"
         Caption         =   "apelido"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "cep"
         Caption         =   "cep"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "endereco"
         Caption         =   "endereco"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "ie"
         Caption         =   "ie"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "fones"
         Caption         =   "fones"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "fax"
         Caption         =   "fax"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "contatos"
         Caption         =   "contatos"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "tiposub"
         Caption         =   "tiposub"
      EndProperty
      BeginProperty Field15 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "apelido"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset64 
      CommandName     =   "Sel_CadProjetos"
      CommDispId      =   1492
      RsDispId        =   1501
      CommandText     =   "select * from tb_cadprojetos where status = '1' order by sequencia"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   6
      BeginProperty Field1 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "sequencia"
         Caption         =   "sequencia"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "projeto"
         Caption         =   "projeto"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field5 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset65 
      CommandName     =   "Sel_CadFilial"
      CommDispId      =   1502
      RsDispId        =   1507
      CommandText     =   "select * from tb_filial where filial = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   6
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "filial"
         Caption         =   "filial"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "nomefilial"
         Caption         =   "nomefilial"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "endereco"
         Caption         =   "endereco"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "telefones"
         Caption         =   "telefones"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Filial"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   50
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset66 
      CommandName     =   "Sel_TempNF"
      CommDispId      =   1508
      RsDispId        =   1921
      CommandText     =   "select * from tb_tempnf where ctrnf = ? order by data"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ctrnf"
         Caption         =   "ctrnf"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "numnf"
         Caption         =   "numnf"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "serie"
         Caption         =   "serie"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valornf"
         Caption         =   "valornf"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesonf"
         Caption         =   "pesonf"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "volumesnf"
         Caption         =   "volumesnf"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "ctr"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset67 
      CommandName     =   "Ins_TempNF"
      CommDispId      =   1512
      RsDispId        =   -1
      CommandText     =   "insert into tb_tempnf (ctrnf, numnf, serie, valornf, pesonf, volumesnf) values ( ? , ? , ? , ? , ? , ? )"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   6
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "ctr"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "numnf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "serie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "valor"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "peso"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "volumes"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset68 
      CommandName     =   "Sel_TempNfTotais"
      CommDispId      =   1514
      RsDispId        =   1568
      CommandText     =   "select sum(valornf) valort, sum(pesonf) pesot, sum(volumesnf) volt, count(*) qtd from tb_tempnf where ctrnf = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valort"
         Caption         =   "valort"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesot"
         Caption         =   "pesot"
      EndProperty
      BeginProperty Field3 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "volt"
         Caption         =   "volt"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "qtd"
         Caption         =   "qtd"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset69 
      CommandName     =   "Sel_CadCliProdsTAB"
      CommDispId      =   1557
      RsDispId        =   1563
      CommandText     =   "select * from tb_cadcliprodstab where status = '1' and cgc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "remetente"
         Caption         =   "remetente"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "nomerem"
         Caption         =   "nomerem"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "natproduto"
         Caption         =   "natproduto"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "tabelapreco"
         Caption         =   "tabelapreco"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricaotab"
         Caption         =   "descricaotab"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "modal"
         Caption         =   "modal"
      EndProperty
      BeginProperty Field9 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset70 
      CommandName     =   "Ins_CadCliProdTAB"
      CommDispId      =   1569
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":14A5
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   10
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "status"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "remetcgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "remetnome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "natproduto"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "tabela"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         UserName        =   "descrtab"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   80
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         UserName        =   "datacad"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         UserName        =   "usuario"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         UserName        =   "modal"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset71 
      CommandName     =   "Sel_TabPrecoEmissao"
      CommDispId      =   1639
      RsDispId        =   1650
      CommandText     =   "select * from tb_cadcliprodstab where status = '1' and cgc = ? and remetente = ? and natproduto = ? and (modal = ? or modal = 'G')"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "remetente"
         Caption         =   "remetente"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "nomerem"
         Caption         =   "nomerem"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "natproduto"
         Caption         =   "natproduto"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "tabelapreco"
         Caption         =   "tabelapreco"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricaotab"
         Caption         =   "descricaotab"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "modal"
         Caption         =   "modal"
      EndProperty
      BeginProperty Field9 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   4
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Consig14"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "Remet8"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "NatProd"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "Modal1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset72 
      CommandName     =   "Sel_CadSubContraNomeLike"
      CommDispId      =   1651
      RsDispId        =   1660
      CommandText     =   "select * from tb_cadsubcontra where nome like ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   17
      BeginProperty Field1 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "nome"
         Caption         =   "nome"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "fantasia"
         Caption         =   "fantasia"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "apelido"
         Caption         =   "apelido"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "cep"
         Caption         =   "cep"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "endereco"
         Caption         =   "endereco"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "ie"
         Caption         =   "ie"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "fones"
         Caption         =   "fones"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "fax"
         Caption         =   "fax"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "contatos"
         Caption         =   "contatos"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "tiposub"
         Caption         =   "tiposub"
      EndProperty
      BeginProperty Field15 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset73 
      CommandName     =   "Sel_TR01UFCidade"
      CommDispId      =   1658
      RsDispId        =   1671
      CommandText     =   "select * from tb_tr01 where status = '1' and codigo = ? and uf = ? and cidade = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   15
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "cim"
         Caption         =   "cim"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretemin"
         Caption         =   "fretemin"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "tarifaperc"
         Caption         =   "tarifaperc"
      EndProperty
      BeginProperty Field12 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field13 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field15 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "UF"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "Cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset74 
      CommandName     =   "Sel_TR01UFCim"
      CommDispId      =   1662
      RsDispId        =   1674
      CommandText     =   "select * from tb_tr01 where status = '1' and codigo = ? and uf = ? and cim = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   15
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "cim"
         Caption         =   "cim"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretemin"
         Caption         =   "fretemin"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "tarifaperc"
         Caption         =   "tarifaperc"
      EndProperty
      BeginProperty Field12 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field13 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field15 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "cim"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset75 
      CommandName     =   "Sel_TA01Sigla"
      CommDispId      =   1675
      RsDispId        =   1680
      CommandText     =   "select * from tb_ta01 where status = '1' and codigo = ? and sigla = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   26
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "localidade"
         Caption         =   "localidade"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "sigla"
         Caption         =   "sigla"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txminima"
         Caption         =   "txminima"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "porkilo"
         Caption         =   "porkilo"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_advalorem"
         Caption         =   "gen_advalorem"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaate"
         Caption         =   "gen_txcoletaate"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletavalor"
         Caption         =   "gen_txcoletavalor"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaexced"
         Caption         =   "gen_txcoletaexced"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaate"
         Caption         =   "gen_txentregaate"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregavalor"
         Caption         =   "gen_txentregavalor"
      EndProperty
      BeginProperty Field18 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaexced"
         Caption         =   "gen_txentregaexced"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      BeginProperty Field20 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_ate"
         Caption         =   "txredesp_ate"
      EndProperty
      BeginProperty Field21 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_valor"
         Caption         =   "txredesp_valor"
      EndProperty
      BeginProperty Field22 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_exced"
         Caption         =   "txredesp_exced"
      EndProperty
      BeginProperty Field23 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field24 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field26 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "Sigla"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset76 
      CommandName     =   "Sel_TG01CodigoAtiva"
      CommDispId      =   1709
      RsDispId        =   1714
      CommandText     =   "select * from tb_tg01 where status = '1' and codigo = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   13
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepeso"
         Caption         =   "fretepeso"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "freteminimo"
         Caption         =   "freteminimo"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalor"
         Caption         =   "fretevalor"
      EndProperty
      BeginProperty Field10 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field13 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset77 
      CommandName     =   "Sel_TempNFNFSerie"
      CommDispId      =   1715
      RsDispId        =   1720
      CommandText     =   "select * from tb_tempnf where ctrnf = ? and numnf = ? and serie = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   6
      BeginProperty Field1 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ctrnf"
         Caption         =   "ctrnf"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "numnf"
         Caption         =   "numnf"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "serie"
         Caption         =   "serie"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valornf"
         Caption         =   "valornf"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesonf"
         Caption         =   "pesonf"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "volumesnf"
         Caption         =   "volumesnf"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Ctr"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "NF"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "Serie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset78 
      CommandName     =   "Sel_TR01Clonar"
      CommDispId      =   1721
      RsDispId        =   1727
      CommandText     =   "select * from tb_tr01 where status = '1' and codigo = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   15
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "cim"
         Caption         =   "cim"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretemin"
         Caption         =   "fretemin"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "tarifaperc"
         Caption         =   "tarifaperc"
      EndProperty
      BeginProperty Field12 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field13 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field15 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset79 
      CommandName     =   "Sel_TA01Clonar"
      CommDispId      =   1728
      RsDispId        =   1733
      CommandText     =   "select * from tb_ta01 where status = '1' and codigo = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   26
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "localidade"
         Caption         =   "localidade"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "sigla"
         Caption         =   "sigla"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txminima"
         Caption         =   "txminima"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "porkilo"
         Caption         =   "porkilo"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_advalorem"
         Caption         =   "gen_advalorem"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaate"
         Caption         =   "gen_txcoletaate"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletavalor"
         Caption         =   "gen_txcoletavalor"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaexced"
         Caption         =   "gen_txcoletaexced"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaate"
         Caption         =   "gen_txentregaate"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregavalor"
         Caption         =   "gen_txentregavalor"
      EndProperty
      BeginProperty Field18 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaexced"
         Caption         =   "gen_txentregaexced"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      BeginProperty Field20 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_ate"
         Caption         =   "txredesp_ate"
      EndProperty
      BeginProperty Field21 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_valor"
         Caption         =   "txredesp_valor"
      EndProperty
      BeginProperty Field22 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredesp_exced"
         Caption         =   "txredesp_exced"
      EndProperty
      BeginProperty Field23 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field24 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field26 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset80 
      CommandName     =   "Sel_TA02Clonar"
      CommDispId      =   1734
      RsDispId        =   1741
      CommandText     =   "select * from tb_ta02 where status = '1' and codigo = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   22
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesode"
         Caption         =   "pesode"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesoate"
         Caption         =   "pesoate"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valormin"
         Caption         =   "valormin"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "porkilo"
         Caption         =   "porkilo"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_advalorem"
         Caption         =   "gen_advalorem"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaate"
         Caption         =   "gen_txcoletaate"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletavalor"
         Caption         =   "gen_txcoletavalor"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaexced"
         Caption         =   "gen_txcoletaexced"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaate"
         Caption         =   "gen_txentregaate"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregavalor"
         Caption         =   "gen_txentregavalor"
      EndProperty
      BeginProperty Field18 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaexced"
         Caption         =   "gen_txentregaexced"
      EndProperty
      BeginProperty Field19 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field20 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field22 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset81 
      CommandName     =   "Sel_TA02PorPeso"
      CommDispId      =   1782
      RsDispId        =   1787
      CommandText     =   "select * from tb_ta02 where status = '1' and codigo = ? and pesode <= ? and pesoate >= ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   22
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesode"
         Caption         =   "pesode"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesoate"
         Caption         =   "pesoate"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valormin"
         Caption         =   "valormin"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "porkilo"
         Caption         =   "porkilo"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_advalorem"
         Caption         =   "gen_advalorem"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaate"
         Caption         =   "gen_txcoletaate"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletavalor"
         Caption         =   "gen_txcoletavalor"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaexced"
         Caption         =   "gen_txcoletaexced"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaate"
         Caption         =   "gen_txentregaate"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregavalor"
         Caption         =   "gen_txentregavalor"
      EndProperty
      BeginProperty Field18 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaexced"
         Caption         =   "gen_txentregaexced"
      EndProperty
      BeginProperty Field19 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field20 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field22 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "Pesode"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "PesoAte"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset82 
      CommandName     =   "Sel_VeiculosCod"
      CommDispId      =   1854
      RsDispId        =   1860
      CommandText     =   "select * from tb_cadveiculos where codigo = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   16
      BeginProperty Field1 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "legenda"
         Caption         =   "legenda"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "frota"
         Caption         =   "frota"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "placa"
         Caption         =   "placa"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "marca"
         Caption         =   "marca"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "modelo"
         Caption         =   "modelo"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   4
         Scale           =   0
         Type            =   200
         Name            =   "ano"
         Caption         =   "ano"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "tipo"
         Caption         =   "tipo"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "suspensaoar"
         Caption         =   "suspensaoar"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "plataformahidr"
         Caption         =   "plataformahidr"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "rastreamento"
         Caption         =   "rastreamento"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "camarafria"
         Caption         =   "camarafria"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "proprietario"
         Caption         =   "proprietario"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "obs"
         Caption         =   "obs"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "codigo"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset83 
      CommandName     =   "Sel_VeiculosPlaca"
      CommDispId      =   1861
      RsDispId        =   1866
      CommandText     =   "select * from tb_cadveiculos where placa = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   18
      BeginProperty Field1 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "legenda"
         Caption         =   "legenda"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "frota"
         Caption         =   "frota"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "placa"
         Caption         =   "placa"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "marca"
         Caption         =   "marca"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "modelo"
         Caption         =   "modelo"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   4
         Scale           =   0
         Type            =   200
         Name            =   "ano"
         Caption         =   "ano"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "tipo"
         Caption         =   "tipo"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "suspensaoar"
         Caption         =   "suspensaoar"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "plataformahidr"
         Caption         =   "plataformahidr"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "rastreamento"
         Caption         =   "rastreamento"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "camarafria"
         Caption         =   "camarafria"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "proprietario"
         Caption         =   "proprietario"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "obs"
         Caption         =   "obs"
      EndProperty
      BeginProperty Field17 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "capacidadepeso"
         Caption         =   "capacidadepeso"
      EndProperty
      BeginProperty Field18 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "capacidadem3"
         Caption         =   "capacidadem3"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "placa"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset84 
      CommandName     =   "Ins_VeicPrestador"
      CommDispId      =   1867
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":155C
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "placa"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "propriet"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset85 
      CommandName     =   "Sel_MotoristaNome"
      CommDispId      =   1869
      RsDispId        =   1874
      CommandText     =   "select * from tb_cadmotoristas where nome like ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "NOME"
         Caption         =   "NOME"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "FUNCAO"
         Caption         =   "FUNCAO"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   11
         Scale           =   0
         Type            =   200
         Name            =   "CPF"
         Caption         =   "CPF"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   11
         Scale           =   0
         Type            =   200
         Name            =   "CNH"
         Caption         =   "CNH"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset86 
      CommandName     =   "Sel_MinutaCTC"
      CommDispId      =   1875
      RsDispId        =   1881
      CommandText     =   "select * from tb_ctc_esp where filialctc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   91
      BeginProperty Field1 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "tipodoc"
         Caption         =   "tipodoc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "motivodoc"
         Caption         =   "motivodoc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "filial"
         Caption         =   "filial"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "ctc"
         Caption         =   "ctc"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hora"
         Caption         =   "hora"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "emissor"
         Caption         =   "emissor"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "prioridade"
         Caption         =   "prioridade"
      EndProperty
      BeginProperty Field10 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "prev_entrega"
         Caption         =   "prev_entrega"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "remet_cgc"
         Caption         =   "remet_cgc"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_nome"
         Caption         =   "remet_nome"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_end"
         Caption         =   "remet_end"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "remet_cidade"
         Caption         =   "remet_cidade"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "remet_uf"
         Caption         =   "remet_uf"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "remet_cep"
         Caption         =   "remet_cep"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "respons_cgc"
         Caption         =   "respons_cgc"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "respons_nome"
         Caption         =   "respons_nome"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "dest_cgc"
         Caption         =   "dest_cgc"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_nome"
         Caption         =   "dest_nome"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_end"
         Caption         =   "dest_end"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "dest_cidade"
         Caption         =   "dest_cidade"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "dest_uf"
         Caption         =   "dest_uf"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "dest_cep"
         Caption         =   "dest_cep"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "cidade_orig"
         Caption         =   "cidade_orig"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_orig"
         Caption         =   "uf_orig"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "via"
         Caption         =   "via"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "cidade_dest"
         Caption         =   "cidade_dest"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_dest"
         Caption         =   "uf_dest"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiao"
         Caption         =   "regiao"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "regiaosac"
         Caption         =   "regiaosac"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "atendsac"
         Caption         =   "atendsac"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "nfs"
         Caption         =   "nfs"
      EndProperty
      BeginProperty Field35 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valmerc"
         Caption         =   "valmerc"
      EndProperty
      BeginProperty Field36 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "peso"
         Caption         =   "peso"
      EndProperty
      BeginProperty Field37 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesotax"
         Caption         =   "pesotax"
      EndProperty
      BeginProperty Field38 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "volumes"
         Caption         =   "volumes"
      EndProperty
      BeginProperty Field39 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "especie"
         Caption         =   "especie"
      EndProperty
      BeginProperty Field40 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "natureza"
         Caption         =   "natureza"
      EndProperty
      BeginProperty Field41 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "naturezaobs"
         Caption         =   "naturezaobs"
      EndProperty
      BeginProperty Field42 
         Precision       =   0
         Size            =   29
         Scale           =   0
         Type            =   200
         Name            =   "dimensoes"
         Caption         =   "dimensoes"
      EndProperty
      BeginProperty Field43 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "tabfrete"
         Caption         =   "tabfrete"
      EndProperty
      BeginProperty Field44 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "tabfretedescr"
         Caption         =   "tabfretedescr"
      EndProperty
      BeginProperty Field45 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepeso"
         Caption         =   "fretepeso"
      EndProperty
      BeginProperty Field46 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalor"
         Caption         =   "fretevalor"
      EndProperty
      BeginProperty Field47 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gris"
         Caption         =   "gris"
      EndProperty
      BeginProperty Field48 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txentregared"
         Caption         =   "txentregared"
      EndProperty
      BeginProperty Field49 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txurgencia"
         Caption         =   "txurgencia"
      EndProperty
      BeginProperty Field50 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pedagio"
         Caption         =   "pedagio"
      EndProperty
      BeginProperty Field51 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretenacional"
         Caption         =   "fretenacional"
      EndProperty
      BeginProperty Field52 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "advalorem"
         Caption         =   "advalorem"
      EndProperty
      BeginProperty Field53 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txorigem"
         Caption         =   "txorigem"
      EndProperty
      BeginProperty Field54 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txdestino"
         Caption         =   "txdestino"
      EndProperty
      BeginProperty Field55 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredespacho"
         Caption         =   "txredespacho"
      EndProperty
      BeginProperty Field56 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txcoleta"
         Caption         =   "txcoleta"
      EndProperty
      BeginProperty Field57 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txoutros"
         Caption         =   "txoutros"
      EndProperty
      BeginProperty Field58 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "descrtxoutros"
         Caption         =   "descrtxoutros"
      EndProperty
      BeginProperty Field59 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotal"
         Caption         =   "fretetotal"
      EndProperty
      BeginProperty Field60 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotalbruto"
         Caption         =   "fretetotalbruto"
      EndProperty
      BeginProperty Field61 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepagosub"
         Caption         =   "fretepagosub"
      EndProperty
      BeginProperty Field62 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tribut"
         Caption         =   "tribut"
      EndProperty
      BeginProperty Field63 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "aliquota"
         Caption         =   "aliquota"
      EndProperty
      BeginProperty Field64 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "filialmanifesto"
         Caption         =   "filialmanifesto"
      EndProperty
      BeginProperty Field65 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "fpag"
         Caption         =   "fpag"
      EndProperty
      BeginProperty Field66 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "faturar"
         Caption         =   "faturar"
      EndProperty
      BeginProperty Field67 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "faturanum"
         Caption         =   "faturanum"
      EndProperty
      BeginProperty Field68 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "entrega_data"
         Caption         =   "entrega_data"
      EndProperty
      BeginProperty Field69 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "recebedor"
         Caption         =   "recebedor"
      EndProperty
      BeginProperty Field70 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   129
         Name            =   "entrega_hora"
         Caption         =   "entrega_hora"
      EndProperty
      BeginProperty Field71 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "modal"
         Caption         =   "modal"
      EndProperty
      BeginProperty Field72 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "obs_emissao"
         Caption         =   "obs_emissao"
      EndProperty
      BeginProperty Field73 
         Precision       =   0
         Size            =   186
         Scale           =   0
         Type            =   200
         Name            =   "obs_ocorr"
         Caption         =   "obs_ocorr"
      EndProperty
      BeginProperty Field74 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "arq"
         Caption         =   "arq"
      EndProperty
      BeginProperty Field75 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tem_ocorr"
         Caption         =   "tem_ocorr"
      EndProperty
      BeginProperty Field76 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "redesp_cgc"
         Caption         =   "redesp_cgc"
      EndProperty
      BeginProperty Field77 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "redesp_nome"
         Caption         =   "redesp_nome"
      EndProperty
      BeginProperty Field78 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "transp_sub"
         Caption         =   "transp_sub"
      EndProperty
      BeginProperty Field79 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "cancelado"
         Caption         =   "cancelado"
      EndProperty
      BeginProperty Field80 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canc_data"
         Caption         =   "canc_data"
      EndProperty
      BeginProperty Field81 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "canc_usu"
         Caption         =   "canc_usu"
      EndProperty
      BeginProperty Field82 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "canc_obs"
         Caption         =   "canc_obs"
      EndProperty
      BeginProperty Field83 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_ctc_cif"
         Caption         =   "at_ctc_cif"
      EndProperty
      BeginProperty Field84 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_ctc_cif_data"
         Caption         =   "at_ctc_cif_data"
      EndProperty
      BeginProperty Field85 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "codevo"
         Caption         =   "codevo"
      EndProperty
      BeginProperty Field86 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_devol"
         Caption         =   "at_devol"
      EndProperty
      BeginProperty Field87 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_cliente"
         Caption         =   "at_cliente"
      EndProperty
      BeginProperty Field88 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_edi"
         Caption         =   "at_edi"
      EndProperty
      BeginProperty Field89 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_edi_data"
         Caption         =   "at_edi_data"
      EndProperty
      BeginProperty Field90 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_edisub"
         Caption         =   "at_edisub"
      EndProperty
      BeginProperty Field91 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_edisub_data"
         Caption         =   "at_edisub_data"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset87 
      CommandName     =   "Ins_TempCTCManifesto"
      CommDispId      =   1882
      RsDispId        =   -1
      CommandText     =   "insert into tb_tempctc (ctrctc, filialctc, data) values ( ? , ? , getdate())"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "ctr"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset88 
      CommandName     =   "Sel_TempMinCTCManif"
      CommDispId      =   1892
      RsDispId        =   2291
      CommandText     =   $"de_informaEM.dsx":1686
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field2 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_nome"
         Caption         =   "remet_nome"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "peso"
         Caption         =   "peso"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "volumes"
         Caption         =   "volumes"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotal"
         Caption         =   "fretetotal"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_nome"
         Caption         =   "dest_nome"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade_dest"
         Caption         =   "cidade_dest"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_dest"
         Caption         =   "uf_dest"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   300
         Scale           =   0
         Type            =   200
         Name            =   "nfs"
         Caption         =   "nfs"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtordem"
         Caption         =   "dtordem"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset89 
      CommandName     =   "Sel_TempMinCTCManifTotal"
      CommDispId      =   1907
      RsDispId        =   1912
      CommandText     =   $"de_informaEM.dsx":1777
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "qtd"
         Caption         =   "qtd"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "tpeso"
         Caption         =   "tpeso"
      EndProperty
      BeginProperty Field3 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "tvol"
         Caption         =   "tvol"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "tval"
         Caption         =   "tval"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "ctr"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset90 
      CommandName     =   "Sel_TempCTCNumero"
      CommDispId      =   1914
      RsDispId        =   1920
      CommandText     =   "select * from tb_tempctc where filialctc = ? and ctrctc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ctrctc"
         Caption         =   "ctrctc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field3 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "ctr"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset91 
      CommandName     =   "Exc_TempCTCMinManif"
      CommDispId      =   1926
      RsDispId        =   -1
      CommandText     =   "delete from tb_tempctc where filialctc = ? and ctrctc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "ctr"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset92 
      CommandName     =   "Exc_TempCTCMinManifTudo"
      CommDispId      =   1928
      RsDispId        =   -1
      CommandText     =   "delete from tb_tempctc where ctrctc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "ctr"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset93 
      CommandName     =   "Ins_TempDimensoes"
      CommDispId      =   1930
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":1830
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   6
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "ctr"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "volpall"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   12
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "qtde"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "largura"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "comprim"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "altura"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset94 
      CommandName     =   "Exc_TempDimensoes"
      CommDispId      =   1939
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":18C0
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   6
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   12
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset95 
      CommandName     =   "Sel_TempDimensoes"
      CommDispId      =   1947
      RsDispId        =   1968
      CommandText     =   "select * from tb_tempdimensoes where ctr = ? order by data"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ctr"
         Caption         =   "ctr"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "volpallet"
         Caption         =   "volpallet"
      EndProperty
      BeginProperty Field3 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "qtdevolpall"
         Caption         =   "qtdevolpall"
      EndProperty
      BeginProperty Field4 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "largura"
         Caption         =   "largura"
      EndProperty
      BeginProperty Field5 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "comprimento"
         Caption         =   "comprimento"
      EndProperty
      BeginProperty Field6 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "altura"
         Caption         =   "altura"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "ctr"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset96 
      CommandName     =   "Ins_MinutaCompleta"
      CommDispId      =   1972
      RsDispId        =   -1
      CommandText     =   "dbo.sp_ins_minutacompl"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CallSyntax      =   $"de_informaEM.dsx":1947
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   61
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@tipodoc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@motivodoc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@filial"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@ctc"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@data"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@hora"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@prioridade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@prev_entrega"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@remet_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "@remet_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "@remet_end"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "@remet_cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "@remet_uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "@remet_cep"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "@respons_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "@respons_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "@dest_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "@dest_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "@dest_end"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P22 
         RealName        =   "@dest_cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P23 
         RealName        =   "@dest_uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P24 
         RealName        =   "@dest_cep"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P25 
         RealName        =   "@cidade_orig"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P26 
         RealName        =   "@uf_orig"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P27 
         RealName        =   "@via"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P28 
         RealName        =   "@cidade_dest"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P29 
         RealName        =   "@uf_dest"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P30 
         RealName        =   "@regiao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P31 
         RealName        =   "@regiaogeo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P32 
         RealName        =   "@regiaosac"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P33 
         RealName        =   "@atendsac"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P34 
         RealName        =   "@nfs"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   300
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P35 
         RealName        =   "@valmerc"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P36 
         RealName        =   "@peso"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P37 
         RealName        =   "@pesotax"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P38 
         RealName        =   "@volumes"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P39 
         RealName        =   "@especie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P40 
         RealName        =   "@natureza"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   25
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P41 
         RealName        =   "@naturezaobs"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   25
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P42 
         RealName        =   "@tabfrete"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P43 
         RealName        =   "@tabfretedescr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   100
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P44 
         RealName        =   "@fretepeso"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P45 
         RealName        =   "@fretevalor"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P46 
         RealName        =   "@gris"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P47 
         RealName        =   "@txcoleta"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P48 
         RealName        =   "@txentregared"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P49 
         RealName        =   "@txurgencia"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P50 
         RealName        =   "@pedagio"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P51 
         RealName        =   "@txoutros"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P52 
         RealName        =   "@descrtxoutros"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P53 
         RealName        =   "@fretetotal"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P54 
         RealName        =   "@fretetotalbruto"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P55 
         RealName        =   "@modal"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P56 
         RealName        =   "@obs_emissao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   320
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P57 
         RealName        =   "@fpag"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P58 
         RealName        =   "@emissor"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P59 
         RealName        =   "@status"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P60 
         RealName        =   "@redesp_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P61 
         RealName        =   "@redesp_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset97 
      CommandName     =   "Ins_NotasFiscais"
      CommDispId      =   1974
      RsDispId        =   -1
      CommandText     =   "dbo.sp_ins_nf"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CallSyntax      =   "{? = CALL dbo.sp_ins_nf( ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   10
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   12
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@numnfnum"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   0
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@numnf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   12
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@serie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@cliente_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@cliente_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@valornf"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@pesonf"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@volumesnf"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset98 
      CommandName     =   "Exc_TempNFCtr"
      CommDispId      =   1979
      RsDispId        =   -1
      CommandText     =   "delete from tb_tempnf where ctrnf = ? "
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "ctr"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset99 
      CommandName     =   "Ins_Dimensoes"
      CommDispId      =   1991
      RsDispId        =   -1
      CommandText     =   "insert into tb_dimensoes (filialctc, volpallet, qtdevolpall, largura, altura, comprimento) values ( ? , ? , ? , ? , ? , ? )"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   6
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "volpallet"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   12
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "qtdevol"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "largura"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "altura"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "comprimento"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset100 
      CommandName     =   "Exc_TempDimensoesCtr"
      CommDispId      =   2004
      RsDispId        =   -1
      CommandText     =   "delete from tb_tempdimensoes where ctr = ? "
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Ctr"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset101 
      CommandName     =   "Exc_TempNfNF"
      CommDispId      =   2006
      RsDispId        =   -1
      CommandText     =   "delete from tb_tempnf where ctrnf = ? and numnf = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "ctr"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "nf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset102 
      CommandName     =   "Sel_Feriado"
      CommDispId      =   2008
      RsDispId        =   2014
      CommandText     =   $"de_informaEM.dsx":1A22
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field2 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ano"
         Caption         =   "ano"
      EndProperty
      BeginProperty Field3 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "mes"
         Caption         =   "mes"
      EndProperty
      BeginProperty Field4 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "dia"
         Caption         =   "dia"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tipo"
         Caption         =   "tipo"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "mes"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "dia"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset103 
      CommandName     =   "Sel_PrazoUF"
      CommDispId      =   2015
      RsDispId        =   2023
      CommandText     =   "select * from tb_cadprazo where codigo = ? and modal = ? and uf = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   6
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "modal"
         Caption         =   "modal"
      EndProperty
      BeginProperty Field4 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "prazo_cap"
         Caption         =   "prazo_cap"
      EndProperty
      BeginProperty Field5 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "prazo_int"
         Caption         =   "prazo_int"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   6
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "modal"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset104 
      CommandName     =   "Sel_CadUfRegiao"
      CommDispId      =   2024
      RsDispId        =   2029
      CommandText     =   "select * from tb_caduf where uf = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "regiaosac"
         Caption         =   "regiaosac"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "atendsac"
         Caption         =   "atendsac"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "UF"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset105 
      CommandName     =   "Ins_Manifesto"
      CommDispId      =   2030
      RsDispId        =   -1
      CommandText     =   "dbo.sp_ins_manifestoinforma"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CallSyntax      =   "{? = CALL dbo.sp_ins_manifestoinforma( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   21
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@filialmanifesto"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@filial"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@manifesto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@motivo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@filialdest"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@lacre"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@dtemissao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@hsemissao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@dtsaida"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "@hssaida"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "@codveiculo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "@placaveic"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "@proprietario"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "@motorista"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "@ajudantes"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   60
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "@conferente"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   25
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "@obs"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "@usuariocad"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "@dataordem"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset106 
      CommandName     =   "Sel_Manifesto"
      CommDispId      =   2032
      RsDispId        =   2163
      CommandText     =   "select * from tb_manifesto where filialmanifesto = ? order by dataordem"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   31
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "idcodigo"
         Caption         =   "idcodigo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "filialmanifesto"
         Caption         =   "filialmanifesto"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "filial"
         Caption         =   "filial"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "manifesto"
         Caption         =   "manifesto"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "motivo"
         Caption         =   "motivo"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "filialdest"
         Caption         =   "filialdest"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "lacre"
         Caption         =   "lacre"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "embarcador"
         Caption         =   "embarcador"
      EndProperty
      BeginProperty Field10 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtemissao"
         Caption         =   "dtemissao"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hsemissao"
         Caption         =   "hsemissao"
      EndProperty
      BeginProperty Field12 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtsaida"
         Caption         =   "dtsaida"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hssaida"
         Caption         =   "hssaida"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "codveiculo"
         Caption         =   "codveiculo"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "placaveic"
         Caption         =   "placaveic"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "proprietario"
         Caption         =   "proprietario"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "motorista"
         Caption         =   "motorista"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "ajudantes"
         Caption         =   "ajudantes"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "conferente"
         Caption         =   "conferente"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "obs"
         Caption         =   "obs"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "cidade_orig"
         Caption         =   "cidade_orig"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_manif_cif"
         Caption         =   "at_manif_cif"
      EndProperty
      BeginProperty Field23 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_manif_cif_data"
         Caption         =   "at_manif_cif_data"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "cancelado"
         Caption         =   "cancelado"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "canc_usu"
         Caption         =   "canc_usu"
      EndProperty
      BeginProperty Field26 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canc_data"
         Caption         =   "canc_data"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "canc_obs"
         Caption         =   "canc_obs"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "saida2"
         Caption         =   "saida2"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "saida2_autorizacao"
         Caption         =   "saida2_autorizacao"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field31 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dataordem"
         Caption         =   "dataordem"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "filialmanifesto"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset107 
      CommandName     =   "Alt_Impressoras"
      CommDispId      =   2038
      RsDispId        =   -1
      CommandText     =   "update tb_impressoras set ctr = ?, manifesto = ?, relatorios = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "ctr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   50
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "manifesto"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   50
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "relatorios"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   50
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset108 
      CommandName     =   "Sel_Impressoras"
      CommDispId      =   2040
      RsDispId        =   2043
      CommandText     =   "select * from tb_impressoras"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "ctr"
         Caption         =   "ctr"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "manifesto"
         Caption         =   "manifesto"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "relatorios"
         Caption         =   "relatorios"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "ctc"
         Caption         =   "ctc"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "minuta"
         Caption         =   "minuta"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset109 
      CommandName     =   "Sel_Usuario"
      CommDispId      =   2095
      RsDispId        =   2100
      CommandText     =   "select * from tb_cadusu where usuario = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "Usuario"
         Caption         =   "Usuario"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "senha"
         Caption         =   "senha"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "Nome"
         Caption         =   "Nome"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   25
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field5 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "DataCad"
         Caption         =   "DataCad"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "stringdireitos"
         Caption         =   "stringdireitos"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "expirada"
         Caption         =   "expirada"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "usuario"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset110 
      CommandName     =   "Alt_Senha"
      CommDispId      =   2101
      RsDispId        =   -1
      CommandText     =   "update tb_cadusu set senha = ?, expirada = 'N'  where usuario = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "senha"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "usuario"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset111 
      CommandName     =   "Alt_CancCTC"
      CommDispId      =   2122
      RsDispId        =   -1
      CommandText     =   "update tb_ctc_esp set tem_ocorr = 'C', cancelado = 'X', canc_data = getdate(), canc_usu = ?, canc_obs  = ? where filialctc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   60
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset112 
      CommandName     =   "alt_descancelar"
      CommDispId      =   2123
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":1A5E
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset113 
      CommandName     =   "Alt_NotFisSitlaSIM"
      CommDispId      =   2128
      RsDispId        =   -1
      CommandText     =   "update tb_ctc_esp set edi_nfsitla = 'S', edi_nfsitladata = getdate() where filialctc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset114 
      CommandName     =   "Sel_NotFisSitlaNumNovos"
      CommDispId      =   2174
      RsDispId        =   2270
      CommandText     =   $"de_informaEM.dsx":1B0A
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   123
      BeginProperty Field1 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "tipodoc"
         Caption         =   "tipodoc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "motivodoc"
         Caption         =   "motivodoc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "filial"
         Caption         =   "filial"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "ctc"
         Caption         =   "ctc"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hora"
         Caption         =   "hora"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "emissor"
         Caption         =   "emissor"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "conferente"
         Caption         =   "conferente"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "prioridade"
         Caption         =   "prioridade"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "prev_entrega"
         Caption         =   "prev_entrega"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "remet_cgc"
         Caption         =   "remet_cgc"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_nome"
         Caption         =   "remet_nome"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_end"
         Caption         =   "remet_end"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "remet_cidade"
         Caption         =   "remet_cidade"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "remet_uf"
         Caption         =   "remet_uf"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "remet_cep"
         Caption         =   "remet_cep"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "respons_cgc"
         Caption         =   "respons_cgc"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "respons_nome"
         Caption         =   "respons_nome"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "dest_cgc"
         Caption         =   "dest_cgc"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_nome"
         Caption         =   "dest_nome"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_end"
         Caption         =   "dest_end"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "dest_cidade"
         Caption         =   "dest_cidade"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "dest_uf"
         Caption         =   "dest_uf"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "dest_cep"
         Caption         =   "dest_cep"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade_orig"
         Caption         =   "cidade_orig"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_orig"
         Caption         =   "uf_orig"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "via"
         Caption         =   "via"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade_dest"
         Caption         =   "cidade_dest"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_dest"
         Caption         =   "uf_dest"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiao"
         Caption         =   "regiao"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "regiaosac"
         Caption         =   "regiaosac"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "atendsac"
         Caption         =   "atendsac"
      EndProperty
      BeginProperty Field35 
         Precision       =   0
         Size            =   300
         Scale           =   0
         Type            =   200
         Name            =   "nfs"
         Caption         =   "nfs"
      EndProperty
      BeginProperty Field36 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valmerc"
         Caption         =   "valmerc"
      EndProperty
      BeginProperty Field37 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "peso"
         Caption         =   "peso"
      EndProperty
      BeginProperty Field38 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesotax"
         Caption         =   "pesotax"
      EndProperty
      BeginProperty Field39 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "volumes"
         Caption         =   "volumes"
      EndProperty
      BeginProperty Field40 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "especie"
         Caption         =   "especie"
      EndProperty
      BeginProperty Field41 
         Precision       =   0
         Size            =   25
         Scale           =   0
         Type            =   200
         Name            =   "natureza"
         Caption         =   "natureza"
      EndProperty
      BeginProperty Field42 
         Precision       =   0
         Size            =   25
         Scale           =   0
         Type            =   200
         Name            =   "naturezaobs"
         Caption         =   "naturezaobs"
      EndProperty
      BeginProperty Field43 
         Precision       =   0
         Size            =   29
         Scale           =   0
         Type            =   200
         Name            =   "dimensoes"
         Caption         =   "dimensoes"
      EndProperty
      BeginProperty Field44 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "tabfrete"
         Caption         =   "tabfrete"
      EndProperty
      BeginProperty Field45 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "tabfretedescr"
         Caption         =   "tabfretedescr"
      EndProperty
      BeginProperty Field46 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepeso"
         Caption         =   "fretepeso"
      EndProperty
      BeginProperty Field47 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalor"
         Caption         =   "fretevalor"
      EndProperty
      BeginProperty Field48 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gris"
         Caption         =   "gris"
      EndProperty
      BeginProperty Field49 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txentregared"
         Caption         =   "txentregared"
      EndProperty
      BeginProperty Field50 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txurgencia"
         Caption         =   "txurgencia"
      EndProperty
      BeginProperty Field51 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pedagio"
         Caption         =   "pedagio"
      EndProperty
      BeginProperty Field52 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretenacional"
         Caption         =   "fretenacional"
      EndProperty
      BeginProperty Field53 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "advalorem"
         Caption         =   "advalorem"
      EndProperty
      BeginProperty Field54 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txorigem"
         Caption         =   "txorigem"
      EndProperty
      BeginProperty Field55 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txdestino"
         Caption         =   "txdestino"
      EndProperty
      BeginProperty Field56 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredespacho"
         Caption         =   "txredespacho"
      EndProperty
      BeginProperty Field57 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txcoleta"
         Caption         =   "txcoleta"
      EndProperty
      BeginProperty Field58 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txoutros"
         Caption         =   "txoutros"
      EndProperty
      BeginProperty Field59 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "descrtxoutros"
         Caption         =   "descrtxoutros"
      EndProperty
      BeginProperty Field60 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotal"
         Caption         =   "fretetotal"
      EndProperty
      BeginProperty Field61 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotalbruto"
         Caption         =   "fretetotalbruto"
      EndProperty
      BeginProperty Field62 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepagosub"
         Caption         =   "fretepagosub"
      EndProperty
      BeginProperty Field63 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tribut"
         Caption         =   "tribut"
      EndProperty
      BeginProperty Field64 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "aliquota"
         Caption         =   "aliquota"
      EndProperty
      BeginProperty Field65 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "filialmanifesto"
         Caption         =   "filialmanifesto"
      EndProperty
      BeginProperty Field66 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "fpag"
         Caption         =   "fpag"
      EndProperty
      BeginProperty Field67 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "faturar"
         Caption         =   "faturar"
      EndProperty
      BeginProperty Field68 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "faturanum"
         Caption         =   "faturanum"
      EndProperty
      BeginProperty Field69 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "entrega_data"
         Caption         =   "entrega_data"
      EndProperty
      BeginProperty Field70 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "recebedor"
         Caption         =   "recebedor"
      EndProperty
      BeginProperty Field71 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   129
         Name            =   "entrega_hora"
         Caption         =   "entrega_hora"
      EndProperty
      BeginProperty Field72 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "modal"
         Caption         =   "modal"
      EndProperty
      BeginProperty Field73 
         Precision       =   0
         Size            =   320
         Scale           =   0
         Type            =   200
         Name            =   "obs_emissao"
         Caption         =   "obs_emissao"
      EndProperty
      BeginProperty Field74 
         Precision       =   0
         Size            =   186
         Scale           =   0
         Type            =   200
         Name            =   "obs_ocorr"
         Caption         =   "obs_ocorr"
      EndProperty
      BeginProperty Field75 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "arq"
         Caption         =   "arq"
      EndProperty
      BeginProperty Field76 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tem_ocorr"
         Caption         =   "tem_ocorr"
      EndProperty
      BeginProperty Field77 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "redesp_cgc"
         Caption         =   "redesp_cgc"
      EndProperty
      BeginProperty Field78 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "redesp_nome"
         Caption         =   "redesp_nome"
      EndProperty
      BeginProperty Field79 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "transp_sub"
         Caption         =   "transp_sub"
      EndProperty
      BeginProperty Field80 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "cancelado"
         Caption         =   "cancelado"
      EndProperty
      BeginProperty Field81 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canc_data"
         Caption         =   "canc_data"
      EndProperty
      BeginProperty Field82 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "canc_usu"
         Caption         =   "canc_usu"
      EndProperty
      BeginProperty Field83 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "canc_obs"
         Caption         =   "canc_obs"
      EndProperty
      BeginProperty Field84 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_ctc_cif"
         Caption         =   "at_ctc_cif"
      EndProperty
      BeginProperty Field85 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_ctc_cif_data"
         Caption         =   "at_ctc_cif_data"
      EndProperty
      BeginProperty Field86 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "codevo"
         Caption         =   "codevo"
      EndProperty
      BeginProperty Field87 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_devol"
         Caption         =   "at_devol"
      EndProperty
      BeginProperty Field88 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_cliente"
         Caption         =   "at_cliente"
      EndProperty
      BeginProperty Field89 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_edi"
         Caption         =   "at_edi"
      EndProperty
      BeginProperty Field90 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_edi_data"
         Caption         =   "at_edi_data"
      EndProperty
      BeginProperty Field91 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_edisub"
         Caption         =   "at_edisub"
      EndProperty
      BeginProperty Field92 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_edisub_data"
         Caption         =   "at_edisub_data"
      EndProperty
      BeginProperty Field93 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "impresso"
         Caption         =   "impresso"
      EndProperty
      BeginProperty Field94 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "impresso_data"
         Caption         =   "impresso_data"
      EndProperty
      BeginProperty Field95 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "impresso_usu"
         Caption         =   "impresso_usu"
      EndProperty
      BeginProperty Field96 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "reimpr"
         Caption         =   "reimpr"
      EndProperty
      BeginProperty Field97 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "reimpr_data"
         Caption         =   "reimpr_data"
      EndProperty
      BeginProperty Field98 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "reimpr_usu"
         Caption         =   "reimpr_usu"
      EndProperty
      BeginProperty Field99 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field100 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "edi_nfsitla"
         Caption         =   "edi_nfsitla"
      EndProperty
      BeginProperty Field101 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "edi_nfsitladata"
         Caption         =   "edi_nfsitladata"
      EndProperty
      BeginProperty Field102 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ordem"
         Caption         =   "ordem"
      EndProperty
      BeginProperty Field103 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "idcodigo"
         Caption         =   "idcodigo"
      EndProperty
      BeginProperty Field104 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field105 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numnfnum"
         Caption         =   "numnfnum"
      EndProperty
      BeginProperty Field106 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "numnf"
         Caption         =   "numnf"
      EndProperty
      BeginProperty Field107 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "serie"
         Caption         =   "serie"
      EndProperty
      BeginProperty Field108 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cliente_cgc"
         Caption         =   "cliente_cgc"
      EndProperty
      BeginProperty Field109 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cliente_nome"
         Caption         =   "cliente_nome"
      EndProperty
      BeginProperty Field110 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "emissao_nf"
         Caption         =   "emissao_nf"
      EndProperty
      BeginProperty Field111 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "numpedido"
         Caption         =   "numpedido"
      EndProperty
      BeginProperty Field112 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtpedido"
         Caption         =   "dtpedido"
      EndProperty
      BeginProperty Field113 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valornf"
         Caption         =   "valornf"
      EndProperty
      BeginProperty Field114 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesonf"
         Caption         =   "pesonf"
      EndProperty
      BeginProperty Field115 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "volumesnf"
         Caption         =   "volumesnf"
      EndProperty
      BeginProperty Field116 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data_interface"
         Caption         =   "data_interface"
      EndProperty
      BeginProperty Field117 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hora_interface"
         Caption         =   "hora_interface"
      EndProperty
      BeginProperty Field118 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_cliente"
         Caption         =   "at_cliente"
      EndProperty
      BeginProperty Field119 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "canhotonf"
         Caption         =   "canhotonf"
      EndProperty
      BeginProperty Field120 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "canhotonfprot"
         Caption         =   "canhotonfprot"
      EndProperty
      BeginProperty Field121 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canhotonfdata"
         Caption         =   "canhotonfdata"
      EndProperty
      BeginProperty Field122 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tem_ocorrnf"
         Caption         =   "tem_ocorrnf"
      EndProperty
      BeginProperty Field123 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ordem"
         Caption         =   "ordem"
      EndProperty
      NumGroups       =   0
      ParamCount      =   8
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Filial"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "CGC"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "CTC1"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   4
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "CTC2"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   4
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "Emissor"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "UF"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         UserName        =   "Modal"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         UserName        =   "Redesp"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset115 
      CommandName     =   "Sel_NotFisSitlaDtNovos"
      CommDispId      =   2196
      RsDispId        =   2271
      CommandText     =   $"de_informaEM.dsx":1CAE
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   123
      BeginProperty Field1 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "tipodoc"
         Caption         =   "tipodoc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "motivodoc"
         Caption         =   "motivodoc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "filial"
         Caption         =   "filial"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "ctc"
         Caption         =   "ctc"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hora"
         Caption         =   "hora"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "emissor"
         Caption         =   "emissor"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "conferente"
         Caption         =   "conferente"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "prioridade"
         Caption         =   "prioridade"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "prev_entrega"
         Caption         =   "prev_entrega"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "remet_cgc"
         Caption         =   "remet_cgc"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_nome"
         Caption         =   "remet_nome"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_end"
         Caption         =   "remet_end"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "remet_cidade"
         Caption         =   "remet_cidade"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "remet_uf"
         Caption         =   "remet_uf"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "remet_cep"
         Caption         =   "remet_cep"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "respons_cgc"
         Caption         =   "respons_cgc"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "respons_nome"
         Caption         =   "respons_nome"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "dest_cgc"
         Caption         =   "dest_cgc"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_nome"
         Caption         =   "dest_nome"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_end"
         Caption         =   "dest_end"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "dest_cidade"
         Caption         =   "dest_cidade"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "dest_uf"
         Caption         =   "dest_uf"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "dest_cep"
         Caption         =   "dest_cep"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade_orig"
         Caption         =   "cidade_orig"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_orig"
         Caption         =   "uf_orig"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "via"
         Caption         =   "via"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade_dest"
         Caption         =   "cidade_dest"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_dest"
         Caption         =   "uf_dest"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiao"
         Caption         =   "regiao"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "regiaosac"
         Caption         =   "regiaosac"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "atendsac"
         Caption         =   "atendsac"
      EndProperty
      BeginProperty Field35 
         Precision       =   0
         Size            =   300
         Scale           =   0
         Type            =   200
         Name            =   "nfs"
         Caption         =   "nfs"
      EndProperty
      BeginProperty Field36 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valmerc"
         Caption         =   "valmerc"
      EndProperty
      BeginProperty Field37 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "peso"
         Caption         =   "peso"
      EndProperty
      BeginProperty Field38 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesotax"
         Caption         =   "pesotax"
      EndProperty
      BeginProperty Field39 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "volumes"
         Caption         =   "volumes"
      EndProperty
      BeginProperty Field40 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "especie"
         Caption         =   "especie"
      EndProperty
      BeginProperty Field41 
         Precision       =   0
         Size            =   25
         Scale           =   0
         Type            =   200
         Name            =   "natureza"
         Caption         =   "natureza"
      EndProperty
      BeginProperty Field42 
         Precision       =   0
         Size            =   25
         Scale           =   0
         Type            =   200
         Name            =   "naturezaobs"
         Caption         =   "naturezaobs"
      EndProperty
      BeginProperty Field43 
         Precision       =   0
         Size            =   29
         Scale           =   0
         Type            =   200
         Name            =   "dimensoes"
         Caption         =   "dimensoes"
      EndProperty
      BeginProperty Field44 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "tabfrete"
         Caption         =   "tabfrete"
      EndProperty
      BeginProperty Field45 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "tabfretedescr"
         Caption         =   "tabfretedescr"
      EndProperty
      BeginProperty Field46 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepeso"
         Caption         =   "fretepeso"
      EndProperty
      BeginProperty Field47 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalor"
         Caption         =   "fretevalor"
      EndProperty
      BeginProperty Field48 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gris"
         Caption         =   "gris"
      EndProperty
      BeginProperty Field49 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txentregared"
         Caption         =   "txentregared"
      EndProperty
      BeginProperty Field50 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txurgencia"
         Caption         =   "txurgencia"
      EndProperty
      BeginProperty Field51 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pedagio"
         Caption         =   "pedagio"
      EndProperty
      BeginProperty Field52 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretenacional"
         Caption         =   "fretenacional"
      EndProperty
      BeginProperty Field53 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "advalorem"
         Caption         =   "advalorem"
      EndProperty
      BeginProperty Field54 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txorigem"
         Caption         =   "txorigem"
      EndProperty
      BeginProperty Field55 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txdestino"
         Caption         =   "txdestino"
      EndProperty
      BeginProperty Field56 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredespacho"
         Caption         =   "txredespacho"
      EndProperty
      BeginProperty Field57 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txcoleta"
         Caption         =   "txcoleta"
      EndProperty
      BeginProperty Field58 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txoutros"
         Caption         =   "txoutros"
      EndProperty
      BeginProperty Field59 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "descrtxoutros"
         Caption         =   "descrtxoutros"
      EndProperty
      BeginProperty Field60 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotal"
         Caption         =   "fretetotal"
      EndProperty
      BeginProperty Field61 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotalbruto"
         Caption         =   "fretetotalbruto"
      EndProperty
      BeginProperty Field62 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepagosub"
         Caption         =   "fretepagosub"
      EndProperty
      BeginProperty Field63 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tribut"
         Caption         =   "tribut"
      EndProperty
      BeginProperty Field64 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "aliquota"
         Caption         =   "aliquota"
      EndProperty
      BeginProperty Field65 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "filialmanifesto"
         Caption         =   "filialmanifesto"
      EndProperty
      BeginProperty Field66 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "fpag"
         Caption         =   "fpag"
      EndProperty
      BeginProperty Field67 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "faturar"
         Caption         =   "faturar"
      EndProperty
      BeginProperty Field68 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "faturanum"
         Caption         =   "faturanum"
      EndProperty
      BeginProperty Field69 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "entrega_data"
         Caption         =   "entrega_data"
      EndProperty
      BeginProperty Field70 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "recebedor"
         Caption         =   "recebedor"
      EndProperty
      BeginProperty Field71 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   129
         Name            =   "entrega_hora"
         Caption         =   "entrega_hora"
      EndProperty
      BeginProperty Field72 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "modal"
         Caption         =   "modal"
      EndProperty
      BeginProperty Field73 
         Precision       =   0
         Size            =   320
         Scale           =   0
         Type            =   200
         Name            =   "obs_emissao"
         Caption         =   "obs_emissao"
      EndProperty
      BeginProperty Field74 
         Precision       =   0
         Size            =   186
         Scale           =   0
         Type            =   200
         Name            =   "obs_ocorr"
         Caption         =   "obs_ocorr"
      EndProperty
      BeginProperty Field75 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "arq"
         Caption         =   "arq"
      EndProperty
      BeginProperty Field76 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tem_ocorr"
         Caption         =   "tem_ocorr"
      EndProperty
      BeginProperty Field77 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "redesp_cgc"
         Caption         =   "redesp_cgc"
      EndProperty
      BeginProperty Field78 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "redesp_nome"
         Caption         =   "redesp_nome"
      EndProperty
      BeginProperty Field79 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "transp_sub"
         Caption         =   "transp_sub"
      EndProperty
      BeginProperty Field80 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "cancelado"
         Caption         =   "cancelado"
      EndProperty
      BeginProperty Field81 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canc_data"
         Caption         =   "canc_data"
      EndProperty
      BeginProperty Field82 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "canc_usu"
         Caption         =   "canc_usu"
      EndProperty
      BeginProperty Field83 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "canc_obs"
         Caption         =   "canc_obs"
      EndProperty
      BeginProperty Field84 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_ctc_cif"
         Caption         =   "at_ctc_cif"
      EndProperty
      BeginProperty Field85 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_ctc_cif_data"
         Caption         =   "at_ctc_cif_data"
      EndProperty
      BeginProperty Field86 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "codevo"
         Caption         =   "codevo"
      EndProperty
      BeginProperty Field87 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_devol"
         Caption         =   "at_devol"
      EndProperty
      BeginProperty Field88 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_cliente"
         Caption         =   "at_cliente"
      EndProperty
      BeginProperty Field89 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_edi"
         Caption         =   "at_edi"
      EndProperty
      BeginProperty Field90 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_edi_data"
         Caption         =   "at_edi_data"
      EndProperty
      BeginProperty Field91 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_edisub"
         Caption         =   "at_edisub"
      EndProperty
      BeginProperty Field92 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_edisub_data"
         Caption         =   "at_edisub_data"
      EndProperty
      BeginProperty Field93 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "impresso"
         Caption         =   "impresso"
      EndProperty
      BeginProperty Field94 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "impresso_data"
         Caption         =   "impresso_data"
      EndProperty
      BeginProperty Field95 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "impresso_usu"
         Caption         =   "impresso_usu"
      EndProperty
      BeginProperty Field96 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "reimpr"
         Caption         =   "reimpr"
      EndProperty
      BeginProperty Field97 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "reimpr_data"
         Caption         =   "reimpr_data"
      EndProperty
      BeginProperty Field98 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "reimpr_usu"
         Caption         =   "reimpr_usu"
      EndProperty
      BeginProperty Field99 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field100 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "edi_nfsitla"
         Caption         =   "edi_nfsitla"
      EndProperty
      BeginProperty Field101 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "edi_nfsitladata"
         Caption         =   "edi_nfsitladata"
      EndProperty
      BeginProperty Field102 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ordem"
         Caption         =   "ordem"
      EndProperty
      BeginProperty Field103 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "idcodigo"
         Caption         =   "idcodigo"
      EndProperty
      BeginProperty Field104 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field105 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numnfnum"
         Caption         =   "numnfnum"
      EndProperty
      BeginProperty Field106 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "numnf"
         Caption         =   "numnf"
      EndProperty
      BeginProperty Field107 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "serie"
         Caption         =   "serie"
      EndProperty
      BeginProperty Field108 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cliente_cgc"
         Caption         =   "cliente_cgc"
      EndProperty
      BeginProperty Field109 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cliente_nome"
         Caption         =   "cliente_nome"
      EndProperty
      BeginProperty Field110 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "emissao_nf"
         Caption         =   "emissao_nf"
      EndProperty
      BeginProperty Field111 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "numpedido"
         Caption         =   "numpedido"
      EndProperty
      BeginProperty Field112 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtpedido"
         Caption         =   "dtpedido"
      EndProperty
      BeginProperty Field113 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valornf"
         Caption         =   "valornf"
      EndProperty
      BeginProperty Field114 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesonf"
         Caption         =   "pesonf"
      EndProperty
      BeginProperty Field115 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "volumesnf"
         Caption         =   "volumesnf"
      EndProperty
      BeginProperty Field116 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data_interface"
         Caption         =   "data_interface"
      EndProperty
      BeginProperty Field117 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hora_interface"
         Caption         =   "hora_interface"
      EndProperty
      BeginProperty Field118 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_cliente"
         Caption         =   "at_cliente"
      EndProperty
      BeginProperty Field119 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "canhotonf"
         Caption         =   "canhotonf"
      EndProperty
      BeginProperty Field120 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "canhotonfprot"
         Caption         =   "canhotonfprot"
      EndProperty
      BeginProperty Field121 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canhotonfdata"
         Caption         =   "canhotonfdata"
      EndProperty
      BeginProperty Field122 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tem_ocorrnf"
         Caption         =   "tem_ocorrnf"
      EndProperty
      BeginProperty Field123 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ordem"
         Caption         =   "ordem"
      EndProperty
      NumGroups       =   0
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Filial"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "CGC"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "Data"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "Emissor"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "UF"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "Modal"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         UserName        =   "Redesp"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset116 
      CommandName     =   "Sel_NotFisSitlaDtTudo"
      CommDispId      =   2214
      RsDispId        =   2272
      CommandText     =   $"de_informaEM.dsx":1E39
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   123
      BeginProperty Field1 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "tipodoc"
         Caption         =   "tipodoc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "motivodoc"
         Caption         =   "motivodoc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "filial"
         Caption         =   "filial"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "ctc"
         Caption         =   "ctc"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hora"
         Caption         =   "hora"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "emissor"
         Caption         =   "emissor"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "conferente"
         Caption         =   "conferente"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "prioridade"
         Caption         =   "prioridade"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "prev_entrega"
         Caption         =   "prev_entrega"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "remet_cgc"
         Caption         =   "remet_cgc"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_nome"
         Caption         =   "remet_nome"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_end"
         Caption         =   "remet_end"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "remet_cidade"
         Caption         =   "remet_cidade"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "remet_uf"
         Caption         =   "remet_uf"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "remet_cep"
         Caption         =   "remet_cep"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "respons_cgc"
         Caption         =   "respons_cgc"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "respons_nome"
         Caption         =   "respons_nome"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "dest_cgc"
         Caption         =   "dest_cgc"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_nome"
         Caption         =   "dest_nome"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_end"
         Caption         =   "dest_end"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "dest_cidade"
         Caption         =   "dest_cidade"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "dest_uf"
         Caption         =   "dest_uf"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "dest_cep"
         Caption         =   "dest_cep"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade_orig"
         Caption         =   "cidade_orig"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_orig"
         Caption         =   "uf_orig"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "via"
         Caption         =   "via"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade_dest"
         Caption         =   "cidade_dest"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_dest"
         Caption         =   "uf_dest"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiao"
         Caption         =   "regiao"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "regiaosac"
         Caption         =   "regiaosac"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "atendsac"
         Caption         =   "atendsac"
      EndProperty
      BeginProperty Field35 
         Precision       =   0
         Size            =   300
         Scale           =   0
         Type            =   200
         Name            =   "nfs"
         Caption         =   "nfs"
      EndProperty
      BeginProperty Field36 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valmerc"
         Caption         =   "valmerc"
      EndProperty
      BeginProperty Field37 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "peso"
         Caption         =   "peso"
      EndProperty
      BeginProperty Field38 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesotax"
         Caption         =   "pesotax"
      EndProperty
      BeginProperty Field39 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "volumes"
         Caption         =   "volumes"
      EndProperty
      BeginProperty Field40 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "especie"
         Caption         =   "especie"
      EndProperty
      BeginProperty Field41 
         Precision       =   0
         Size            =   25
         Scale           =   0
         Type            =   200
         Name            =   "natureza"
         Caption         =   "natureza"
      EndProperty
      BeginProperty Field42 
         Precision       =   0
         Size            =   25
         Scale           =   0
         Type            =   200
         Name            =   "naturezaobs"
         Caption         =   "naturezaobs"
      EndProperty
      BeginProperty Field43 
         Precision       =   0
         Size            =   29
         Scale           =   0
         Type            =   200
         Name            =   "dimensoes"
         Caption         =   "dimensoes"
      EndProperty
      BeginProperty Field44 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "tabfrete"
         Caption         =   "tabfrete"
      EndProperty
      BeginProperty Field45 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "tabfretedescr"
         Caption         =   "tabfretedescr"
      EndProperty
      BeginProperty Field46 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepeso"
         Caption         =   "fretepeso"
      EndProperty
      BeginProperty Field47 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalor"
         Caption         =   "fretevalor"
      EndProperty
      BeginProperty Field48 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gris"
         Caption         =   "gris"
      EndProperty
      BeginProperty Field49 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txentregared"
         Caption         =   "txentregared"
      EndProperty
      BeginProperty Field50 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txurgencia"
         Caption         =   "txurgencia"
      EndProperty
      BeginProperty Field51 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pedagio"
         Caption         =   "pedagio"
      EndProperty
      BeginProperty Field52 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretenacional"
         Caption         =   "fretenacional"
      EndProperty
      BeginProperty Field53 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "advalorem"
         Caption         =   "advalorem"
      EndProperty
      BeginProperty Field54 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txorigem"
         Caption         =   "txorigem"
      EndProperty
      BeginProperty Field55 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txdestino"
         Caption         =   "txdestino"
      EndProperty
      BeginProperty Field56 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredespacho"
         Caption         =   "txredespacho"
      EndProperty
      BeginProperty Field57 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txcoleta"
         Caption         =   "txcoleta"
      EndProperty
      BeginProperty Field58 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txoutros"
         Caption         =   "txoutros"
      EndProperty
      BeginProperty Field59 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "descrtxoutros"
         Caption         =   "descrtxoutros"
      EndProperty
      BeginProperty Field60 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotal"
         Caption         =   "fretetotal"
      EndProperty
      BeginProperty Field61 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotalbruto"
         Caption         =   "fretetotalbruto"
      EndProperty
      BeginProperty Field62 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepagosub"
         Caption         =   "fretepagosub"
      EndProperty
      BeginProperty Field63 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tribut"
         Caption         =   "tribut"
      EndProperty
      BeginProperty Field64 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "aliquota"
         Caption         =   "aliquota"
      EndProperty
      BeginProperty Field65 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "filialmanifesto"
         Caption         =   "filialmanifesto"
      EndProperty
      BeginProperty Field66 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "fpag"
         Caption         =   "fpag"
      EndProperty
      BeginProperty Field67 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "faturar"
         Caption         =   "faturar"
      EndProperty
      BeginProperty Field68 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "faturanum"
         Caption         =   "faturanum"
      EndProperty
      BeginProperty Field69 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "entrega_data"
         Caption         =   "entrega_data"
      EndProperty
      BeginProperty Field70 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "recebedor"
         Caption         =   "recebedor"
      EndProperty
      BeginProperty Field71 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   129
         Name            =   "entrega_hora"
         Caption         =   "entrega_hora"
      EndProperty
      BeginProperty Field72 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "modal"
         Caption         =   "modal"
      EndProperty
      BeginProperty Field73 
         Precision       =   0
         Size            =   320
         Scale           =   0
         Type            =   200
         Name            =   "obs_emissao"
         Caption         =   "obs_emissao"
      EndProperty
      BeginProperty Field74 
         Precision       =   0
         Size            =   186
         Scale           =   0
         Type            =   200
         Name            =   "obs_ocorr"
         Caption         =   "obs_ocorr"
      EndProperty
      BeginProperty Field75 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "arq"
         Caption         =   "arq"
      EndProperty
      BeginProperty Field76 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tem_ocorr"
         Caption         =   "tem_ocorr"
      EndProperty
      BeginProperty Field77 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "redesp_cgc"
         Caption         =   "redesp_cgc"
      EndProperty
      BeginProperty Field78 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "redesp_nome"
         Caption         =   "redesp_nome"
      EndProperty
      BeginProperty Field79 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "transp_sub"
         Caption         =   "transp_sub"
      EndProperty
      BeginProperty Field80 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "cancelado"
         Caption         =   "cancelado"
      EndProperty
      BeginProperty Field81 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canc_data"
         Caption         =   "canc_data"
      EndProperty
      BeginProperty Field82 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "canc_usu"
         Caption         =   "canc_usu"
      EndProperty
      BeginProperty Field83 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "canc_obs"
         Caption         =   "canc_obs"
      EndProperty
      BeginProperty Field84 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_ctc_cif"
         Caption         =   "at_ctc_cif"
      EndProperty
      BeginProperty Field85 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_ctc_cif_data"
         Caption         =   "at_ctc_cif_data"
      EndProperty
      BeginProperty Field86 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "codevo"
         Caption         =   "codevo"
      EndProperty
      BeginProperty Field87 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_devol"
         Caption         =   "at_devol"
      EndProperty
      BeginProperty Field88 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_cliente"
         Caption         =   "at_cliente"
      EndProperty
      BeginProperty Field89 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_edi"
         Caption         =   "at_edi"
      EndProperty
      BeginProperty Field90 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_edi_data"
         Caption         =   "at_edi_data"
      EndProperty
      BeginProperty Field91 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_edisub"
         Caption         =   "at_edisub"
      EndProperty
      BeginProperty Field92 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_edisub_data"
         Caption         =   "at_edisub_data"
      EndProperty
      BeginProperty Field93 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "impresso"
         Caption         =   "impresso"
      EndProperty
      BeginProperty Field94 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "impresso_data"
         Caption         =   "impresso_data"
      EndProperty
      BeginProperty Field95 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "impresso_usu"
         Caption         =   "impresso_usu"
      EndProperty
      BeginProperty Field96 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "reimpr"
         Caption         =   "reimpr"
      EndProperty
      BeginProperty Field97 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "reimpr_data"
         Caption         =   "reimpr_data"
      EndProperty
      BeginProperty Field98 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "reimpr_usu"
         Caption         =   "reimpr_usu"
      EndProperty
      BeginProperty Field99 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field100 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "edi_nfsitla"
         Caption         =   "edi_nfsitla"
      EndProperty
      BeginProperty Field101 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "edi_nfsitladata"
         Caption         =   "edi_nfsitladata"
      EndProperty
      BeginProperty Field102 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ordem"
         Caption         =   "ordem"
      EndProperty
      BeginProperty Field103 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "idcodigo"
         Caption         =   "idcodigo"
      EndProperty
      BeginProperty Field104 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field105 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numnfnum"
         Caption         =   "numnfnum"
      EndProperty
      BeginProperty Field106 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "numnf"
         Caption         =   "numnf"
      EndProperty
      BeginProperty Field107 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "serie"
         Caption         =   "serie"
      EndProperty
      BeginProperty Field108 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cliente_cgc"
         Caption         =   "cliente_cgc"
      EndProperty
      BeginProperty Field109 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cliente_nome"
         Caption         =   "cliente_nome"
      EndProperty
      BeginProperty Field110 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "emissao_nf"
         Caption         =   "emissao_nf"
      EndProperty
      BeginProperty Field111 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "numpedido"
         Caption         =   "numpedido"
      EndProperty
      BeginProperty Field112 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtpedido"
         Caption         =   "dtpedido"
      EndProperty
      BeginProperty Field113 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valornf"
         Caption         =   "valornf"
      EndProperty
      BeginProperty Field114 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesonf"
         Caption         =   "pesonf"
      EndProperty
      BeginProperty Field115 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "volumesnf"
         Caption         =   "volumesnf"
      EndProperty
      BeginProperty Field116 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data_interface"
         Caption         =   "data_interface"
      EndProperty
      BeginProperty Field117 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hora_interface"
         Caption         =   "hora_interface"
      EndProperty
      BeginProperty Field118 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_cliente"
         Caption         =   "at_cliente"
      EndProperty
      BeginProperty Field119 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "canhotonf"
         Caption         =   "canhotonf"
      EndProperty
      BeginProperty Field120 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "canhotonfprot"
         Caption         =   "canhotonfprot"
      EndProperty
      BeginProperty Field121 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canhotonfdata"
         Caption         =   "canhotonfdata"
      EndProperty
      BeginProperty Field122 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tem_ocorrnf"
         Caption         =   "tem_ocorrnf"
      EndProperty
      BeginProperty Field123 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ordem"
         Caption         =   "ordem"
      EndProperty
      NumGroups       =   0
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "Filial"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "CGC"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "Data"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "Emissor"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "UF"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "Redesp"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         UserName        =   "Modal"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset117 
      CommandName     =   "Sel_NotFisSitlaNumTudo"
      CommDispId      =   2224
      RsDispId        =   2358
      CommandText     =   $"de_informaEM.dsx":1F8F
      ActiveConnectionName=   "cn_informa"
      CommandTimeout  =   240
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   138
      BeginProperty Field1 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "tipodoc"
         Caption         =   "tipodoc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "motivodoc"
         Caption         =   "motivodoc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "filial"
         Caption         =   "filial"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "ctc"
         Caption         =   "ctc"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hora"
         Caption         =   "hora"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "emissor"
         Caption         =   "emissor"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "conferente"
         Caption         =   "conferente"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "prioridade"
         Caption         =   "prioridade"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "prev_entrega"
         Caption         =   "prev_entrega"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "remet_cgc"
         Caption         =   "remet_cgc"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_nome"
         Caption         =   "remet_nome"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_end"
         Caption         =   "remet_end"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "remet_cidade"
         Caption         =   "remet_cidade"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "remet_uf"
         Caption         =   "remet_uf"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "remet_cep"
         Caption         =   "remet_cep"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "respons_cgc"
         Caption         =   "respons_cgc"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "respons_nome"
         Caption         =   "respons_nome"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "dest_cgc"
         Caption         =   "dest_cgc"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_nome"
         Caption         =   "dest_nome"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_end"
         Caption         =   "dest_end"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "dest_cidade"
         Caption         =   "dest_cidade"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "dest_uf"
         Caption         =   "dest_uf"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "dest_cep"
         Caption         =   "dest_cep"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade_orig"
         Caption         =   "cidade_orig"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_orig"
         Caption         =   "uf_orig"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "via"
         Caption         =   "via"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade_dest"
         Caption         =   "cidade_dest"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_dest"
         Caption         =   "uf_dest"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiao"
         Caption         =   "regiao"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "regiaosac"
         Caption         =   "regiaosac"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "atendsac"
         Caption         =   "atendsac"
      EndProperty
      BeginProperty Field35 
         Precision       =   0
         Size            =   300
         Scale           =   0
         Type            =   200
         Name            =   "nfs"
         Caption         =   "nfs"
      EndProperty
      BeginProperty Field36 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valmerc"
         Caption         =   "valmerc"
      EndProperty
      BeginProperty Field37 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "peso"
         Caption         =   "peso"
      EndProperty
      BeginProperty Field38 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesotax"
         Caption         =   "pesotax"
      EndProperty
      BeginProperty Field39 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "volumes"
         Caption         =   "volumes"
      EndProperty
      BeginProperty Field40 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "especie"
         Caption         =   "especie"
      EndProperty
      BeginProperty Field41 
         Precision       =   0
         Size            =   25
         Scale           =   0
         Type            =   200
         Name            =   "natureza"
         Caption         =   "natureza"
      EndProperty
      BeginProperty Field42 
         Precision       =   0
         Size            =   25
         Scale           =   0
         Type            =   200
         Name            =   "naturezaobs"
         Caption         =   "naturezaobs"
      EndProperty
      BeginProperty Field43 
         Precision       =   0
         Size            =   29
         Scale           =   0
         Type            =   200
         Name            =   "dimensoes"
         Caption         =   "dimensoes"
      EndProperty
      BeginProperty Field44 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "tabfrete"
         Caption         =   "tabfrete"
      EndProperty
      BeginProperty Field45 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "tabfretedescr"
         Caption         =   "tabfretedescr"
      EndProperty
      BeginProperty Field46 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepeso"
         Caption         =   "fretepeso"
      EndProperty
      BeginProperty Field47 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalor"
         Caption         =   "fretevalor"
      EndProperty
      BeginProperty Field48 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gris"
         Caption         =   "gris"
      EndProperty
      BeginProperty Field49 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txentregared"
         Caption         =   "txentregared"
      EndProperty
      BeginProperty Field50 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txurgencia"
         Caption         =   "txurgencia"
      EndProperty
      BeginProperty Field51 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pedagio"
         Caption         =   "pedagio"
      EndProperty
      BeginProperty Field52 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretenacional"
         Caption         =   "fretenacional"
      EndProperty
      BeginProperty Field53 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "advalorem"
         Caption         =   "advalorem"
      EndProperty
      BeginProperty Field54 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txorigem"
         Caption         =   "txorigem"
      EndProperty
      BeginProperty Field55 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txdestino"
         Caption         =   "txdestino"
      EndProperty
      BeginProperty Field56 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredespacho"
         Caption         =   "txredespacho"
      EndProperty
      BeginProperty Field57 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txcoleta"
         Caption         =   "txcoleta"
      EndProperty
      BeginProperty Field58 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txoutros"
         Caption         =   "txoutros"
      EndProperty
      BeginProperty Field59 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "descrtxoutros"
         Caption         =   "descrtxoutros"
      EndProperty
      BeginProperty Field60 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotal"
         Caption         =   "fretetotal"
      EndProperty
      BeginProperty Field61 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotalbruto"
         Caption         =   "fretetotalbruto"
      EndProperty
      BeginProperty Field62 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepagosub"
         Caption         =   "fretepagosub"
      EndProperty
      BeginProperty Field63 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tribut"
         Caption         =   "tribut"
      EndProperty
      BeginProperty Field64 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "aliquota"
         Caption         =   "aliquota"
      EndProperty
      BeginProperty Field65 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "filialmanifesto"
         Caption         =   "filialmanifesto"
      EndProperty
      BeginProperty Field66 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "fpag"
         Caption         =   "fpag"
      EndProperty
      BeginProperty Field67 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "faturar"
         Caption         =   "faturar"
      EndProperty
      BeginProperty Field68 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "faturanum"
         Caption         =   "faturanum"
      EndProperty
      BeginProperty Field69 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "entrega_data"
         Caption         =   "entrega_data"
      EndProperty
      BeginProperty Field70 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "recebedor"
         Caption         =   "recebedor"
      EndProperty
      BeginProperty Field71 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   129
         Name            =   "entrega_hora"
         Caption         =   "entrega_hora"
      EndProperty
      BeginProperty Field72 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "modal"
         Caption         =   "modal"
      EndProperty
      BeginProperty Field73 
         Precision       =   0
         Size            =   320
         Scale           =   0
         Type            =   200
         Name            =   "obs_emissao"
         Caption         =   "obs_emissao"
      EndProperty
      BeginProperty Field74 
         Precision       =   0
         Size            =   186
         Scale           =   0
         Type            =   200
         Name            =   "obs_ocorr"
         Caption         =   "obs_ocorr"
      EndProperty
      BeginProperty Field75 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "arq"
         Caption         =   "arq"
      EndProperty
      BeginProperty Field76 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tem_ocorr"
         Caption         =   "tem_ocorr"
      EndProperty
      BeginProperty Field77 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "redesp_cgc"
         Caption         =   "redesp_cgc"
      EndProperty
      BeginProperty Field78 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "redesp_nome"
         Caption         =   "redesp_nome"
      EndProperty
      BeginProperty Field79 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "transp_sub"
         Caption         =   "transp_sub"
      EndProperty
      BeginProperty Field80 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "cancelado"
         Caption         =   "cancelado"
      EndProperty
      BeginProperty Field81 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canc_data"
         Caption         =   "canc_data"
      EndProperty
      BeginProperty Field82 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "canc_usu"
         Caption         =   "canc_usu"
      EndProperty
      BeginProperty Field83 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "canc_obs"
         Caption         =   "canc_obs"
      EndProperty
      BeginProperty Field84 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_ctc_cif"
         Caption         =   "at_ctc_cif"
      EndProperty
      BeginProperty Field85 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_ctc_cif_data"
         Caption         =   "at_ctc_cif_data"
      EndProperty
      BeginProperty Field86 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "codevo"
         Caption         =   "codevo"
      EndProperty
      BeginProperty Field87 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_devol"
         Caption         =   "at_devol"
      EndProperty
      BeginProperty Field88 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_cliente"
         Caption         =   "at_cliente"
      EndProperty
      BeginProperty Field89 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_edi"
         Caption         =   "at_edi"
      EndProperty
      BeginProperty Field90 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_edi_data"
         Caption         =   "at_edi_data"
      EndProperty
      BeginProperty Field91 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_edisub"
         Caption         =   "at_edisub"
      EndProperty
      BeginProperty Field92 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_edisub_data"
         Caption         =   "at_edisub_data"
      EndProperty
      BeginProperty Field93 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "impresso"
         Caption         =   "impresso"
      EndProperty
      BeginProperty Field94 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "impresso_data"
         Caption         =   "impresso_data"
      EndProperty
      BeginProperty Field95 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "impresso_usu"
         Caption         =   "impresso_usu"
      EndProperty
      BeginProperty Field96 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "reimpr"
         Caption         =   "reimpr"
      EndProperty
      BeginProperty Field97 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "reimpr_data"
         Caption         =   "reimpr_data"
      EndProperty
      BeginProperty Field98 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "reimpr_usu"
         Caption         =   "reimpr_usu"
      EndProperty
      BeginProperty Field99 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field100 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "edi_nfsitla"
         Caption         =   "edi_nfsitla"
      EndProperty
      BeginProperty Field101 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "edi_nfsitladata"
         Caption         =   "edi_nfsitladata"
      EndProperty
      BeginProperty Field102 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ordem"
         Caption         =   "ordem"
      EndProperty
      BeginProperty Field103 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valmercCTR03"
         Caption         =   "valmercCTR03"
      EndProperty
      BeginProperty Field104 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesoCTR03"
         Caption         =   "pesoCTR03"
      EndProperty
      BeginProperty Field105 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesotaxCTR03"
         Caption         =   "pesotaxCTR03"
      EndProperty
      BeginProperty Field106 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "volumesCTR03"
         Caption         =   "volumesCTR03"
      EndProperty
      BeginProperty Field107 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepesoCTR03"
         Caption         =   "fretepesoCTR03"
      EndProperty
      BeginProperty Field108 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalorCTR03"
         Caption         =   "fretevalorCTR03"
      EndProperty
      BeginProperty Field109 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "grisCTR03"
         Caption         =   "grisCTR03"
      EndProperty
      BeginProperty Field110 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txentregaredCTR03"
         Caption         =   "txentregaredCTR03"
      EndProperty
      BeginProperty Field111 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txurgenciaCTR03"
         Caption         =   "txurgenciaCTR03"
      EndProperty
      BeginProperty Field112 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pedagioCTR03"
         Caption         =   "pedagioCTR03"
      EndProperty
      BeginProperty Field113 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txcoletaCTR03"
         Caption         =   "txcoletaCTR03"
      EndProperty
      BeginProperty Field114 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txoutrosCTR03"
         Caption         =   "txoutrosCTR03"
      EndProperty
      BeginProperty Field115 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "descrtxoutrosCTR03"
         Caption         =   "descrtxoutrosCTR03"
      EndProperty
      BeginProperty Field116 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotalCTR03"
         Caption         =   "fretetotalCTR03"
      EndProperty
      BeginProperty Field117 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotalbrutoCTR03"
         Caption         =   "fretetotalbrutoCTR03"
      EndProperty
      BeginProperty Field118 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "idcodigo"
         Caption         =   "idcodigo"
      EndProperty
      BeginProperty Field119 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field120 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numnfnum"
         Caption         =   "numnfnum"
      EndProperty
      BeginProperty Field121 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "numnf"
         Caption         =   "numnf"
      EndProperty
      BeginProperty Field122 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "serie"
         Caption         =   "serie"
      EndProperty
      BeginProperty Field123 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cliente_cgc"
         Caption         =   "cliente_cgc"
      EndProperty
      BeginProperty Field124 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cliente_nome"
         Caption         =   "cliente_nome"
      EndProperty
      BeginProperty Field125 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "emissao_nf"
         Caption         =   "emissao_nf"
      EndProperty
      BeginProperty Field126 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "numpedido"
         Caption         =   "numpedido"
      EndProperty
      BeginProperty Field127 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtpedido"
         Caption         =   "dtpedido"
      EndProperty
      BeginProperty Field128 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valornf"
         Caption         =   "valornf"
      EndProperty
      BeginProperty Field129 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesonf"
         Caption         =   "pesonf"
      EndProperty
      BeginProperty Field130 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "volumesnf"
         Caption         =   "volumesnf"
      EndProperty
      BeginProperty Field131 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data_interface"
         Caption         =   "data_interface"
      EndProperty
      BeginProperty Field132 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hora_interface"
         Caption         =   "hora_interface"
      EndProperty
      BeginProperty Field133 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_cliente"
         Caption         =   "at_cliente"
      EndProperty
      BeginProperty Field134 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "canhotonf"
         Caption         =   "canhotonf"
      EndProperty
      BeginProperty Field135 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "canhotonfprot"
         Caption         =   "canhotonfprot"
      EndProperty
      BeginProperty Field136 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canhotonfdata"
         Caption         =   "canhotonfdata"
      EndProperty
      BeginProperty Field137 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tem_ocorrnf"
         Caption         =   "tem_ocorrnf"
      EndProperty
      BeginProperty Field138 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ordem"
         Caption         =   "ordem"
      EndProperty
      NumGroups       =   0
      ParamCount      =   8
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   4
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   4
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset118 
      CommandName     =   "Sel_BuscaTranspSubAnterior"
      CommDispId      =   2239
      RsDispId        =   2244
      CommandText     =   "select * from tb_transpsub where cgc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   6
      BeginProperty Field1 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "cgc"
         Caption         =   "cgc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "texto"
         Caption         =   "texto"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "transportador"
         Caption         =   "transportador"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "razaosoc"
         Caption         =   "razaosoc"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "percentual"
         Caption         =   "percentual"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "minimo"
         Caption         =   "minimo"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset119 
      CommandName     =   "Ins_TranspSubAnterior"
      CommDispId      =   2245
      RsDispId        =   -1
      CommandText     =   "insert into tb_transpsub (cgc, transportador, razaosoc) values ( ?, ?, ? )"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "transp"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "razao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset120 
      CommandName     =   "Alt_TranspSubMinCTC"
      CommDispId      =   2253
      RsDispId        =   -1
      CommandText     =   "update tb_ctc_esp set transp_sub = ? where filialctc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "transpsub"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset121 
      CommandName     =   "Sel_AcertoTranspSub"
      CommDispId      =   2255
      RsDispId        =   2261
      CommandText     =   "select * from tb_ctc_esp where tipodoc = 'MC' and tem_ocorr <> 'C' and len(transp_sub) < 3 and len(redesp_cgc) > 5"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   99
      BeginProperty Field1 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "tipodoc"
         Caption         =   "tipodoc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "motivodoc"
         Caption         =   "motivodoc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "filial"
         Caption         =   "filial"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "ctc"
         Caption         =   "ctc"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hora"
         Caption         =   "hora"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "emissor"
         Caption         =   "emissor"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "conferente"
         Caption         =   "conferente"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "prioridade"
         Caption         =   "prioridade"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "prev_entrega"
         Caption         =   "prev_entrega"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "remet_cgc"
         Caption         =   "remet_cgc"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_nome"
         Caption         =   "remet_nome"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_end"
         Caption         =   "remet_end"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "remet_cidade"
         Caption         =   "remet_cidade"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "remet_uf"
         Caption         =   "remet_uf"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "remet_cep"
         Caption         =   "remet_cep"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "respons_cgc"
         Caption         =   "respons_cgc"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "respons_nome"
         Caption         =   "respons_nome"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "dest_cgc"
         Caption         =   "dest_cgc"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_nome"
         Caption         =   "dest_nome"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_end"
         Caption         =   "dest_end"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "dest_cidade"
         Caption         =   "dest_cidade"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "dest_uf"
         Caption         =   "dest_uf"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "dest_cep"
         Caption         =   "dest_cep"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade_orig"
         Caption         =   "cidade_orig"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_orig"
         Caption         =   "uf_orig"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "via"
         Caption         =   "via"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade_dest"
         Caption         =   "cidade_dest"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_dest"
         Caption         =   "uf_dest"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiao"
         Caption         =   "regiao"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "regiaosac"
         Caption         =   "regiaosac"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "atendsac"
         Caption         =   "atendsac"
      EndProperty
      BeginProperty Field35 
         Precision       =   0
         Size            =   300
         Scale           =   0
         Type            =   200
         Name            =   "nfs"
         Caption         =   "nfs"
      EndProperty
      BeginProperty Field36 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valmerc"
         Caption         =   "valmerc"
      EndProperty
      BeginProperty Field37 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "peso"
         Caption         =   "peso"
      EndProperty
      BeginProperty Field38 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesotax"
         Caption         =   "pesotax"
      EndProperty
      BeginProperty Field39 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "volumes"
         Caption         =   "volumes"
      EndProperty
      BeginProperty Field40 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "especie"
         Caption         =   "especie"
      EndProperty
      BeginProperty Field41 
         Precision       =   0
         Size            =   25
         Scale           =   0
         Type            =   200
         Name            =   "natureza"
         Caption         =   "natureza"
      EndProperty
      BeginProperty Field42 
         Precision       =   0
         Size            =   25
         Scale           =   0
         Type            =   200
         Name            =   "naturezaobs"
         Caption         =   "naturezaobs"
      EndProperty
      BeginProperty Field43 
         Precision       =   0
         Size            =   29
         Scale           =   0
         Type            =   200
         Name            =   "dimensoes"
         Caption         =   "dimensoes"
      EndProperty
      BeginProperty Field44 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "tabfrete"
         Caption         =   "tabfrete"
      EndProperty
      BeginProperty Field45 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "tabfretedescr"
         Caption         =   "tabfretedescr"
      EndProperty
      BeginProperty Field46 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepeso"
         Caption         =   "fretepeso"
      EndProperty
      BeginProperty Field47 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalor"
         Caption         =   "fretevalor"
      EndProperty
      BeginProperty Field48 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gris"
         Caption         =   "gris"
      EndProperty
      BeginProperty Field49 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txentregared"
         Caption         =   "txentregared"
      EndProperty
      BeginProperty Field50 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txurgencia"
         Caption         =   "txurgencia"
      EndProperty
      BeginProperty Field51 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pedagio"
         Caption         =   "pedagio"
      EndProperty
      BeginProperty Field52 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretenacional"
         Caption         =   "fretenacional"
      EndProperty
      BeginProperty Field53 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "advalorem"
         Caption         =   "advalorem"
      EndProperty
      BeginProperty Field54 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txorigem"
         Caption         =   "txorigem"
      EndProperty
      BeginProperty Field55 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txdestino"
         Caption         =   "txdestino"
      EndProperty
      BeginProperty Field56 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredespacho"
         Caption         =   "txredespacho"
      EndProperty
      BeginProperty Field57 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txcoleta"
         Caption         =   "txcoleta"
      EndProperty
      BeginProperty Field58 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txoutros"
         Caption         =   "txoutros"
      EndProperty
      BeginProperty Field59 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "descrtxoutros"
         Caption         =   "descrtxoutros"
      EndProperty
      BeginProperty Field60 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotal"
         Caption         =   "fretetotal"
      EndProperty
      BeginProperty Field61 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotalbruto"
         Caption         =   "fretetotalbruto"
      EndProperty
      BeginProperty Field62 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepagosub"
         Caption         =   "fretepagosub"
      EndProperty
      BeginProperty Field63 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tribut"
         Caption         =   "tribut"
      EndProperty
      BeginProperty Field64 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "aliquota"
         Caption         =   "aliquota"
      EndProperty
      BeginProperty Field65 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "filialmanifesto"
         Caption         =   "filialmanifesto"
      EndProperty
      BeginProperty Field66 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "fpag"
         Caption         =   "fpag"
      EndProperty
      BeginProperty Field67 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "faturar"
         Caption         =   "faturar"
      EndProperty
      BeginProperty Field68 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "faturanum"
         Caption         =   "faturanum"
      EndProperty
      BeginProperty Field69 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "entrega_data"
         Caption         =   "entrega_data"
      EndProperty
      BeginProperty Field70 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "recebedor"
         Caption         =   "recebedor"
      EndProperty
      BeginProperty Field71 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   129
         Name            =   "entrega_hora"
         Caption         =   "entrega_hora"
      EndProperty
      BeginProperty Field72 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "modal"
         Caption         =   "modal"
      EndProperty
      BeginProperty Field73 
         Precision       =   0
         Size            =   320
         Scale           =   0
         Type            =   200
         Name            =   "obs_emissao"
         Caption         =   "obs_emissao"
      EndProperty
      BeginProperty Field74 
         Precision       =   0
         Size            =   186
         Scale           =   0
         Type            =   200
         Name            =   "obs_ocorr"
         Caption         =   "obs_ocorr"
      EndProperty
      BeginProperty Field75 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "arq"
         Caption         =   "arq"
      EndProperty
      BeginProperty Field76 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tem_ocorr"
         Caption         =   "tem_ocorr"
      EndProperty
      BeginProperty Field77 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "redesp_cgc"
         Caption         =   "redesp_cgc"
      EndProperty
      BeginProperty Field78 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "redesp_nome"
         Caption         =   "redesp_nome"
      EndProperty
      BeginProperty Field79 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "transp_sub"
         Caption         =   "transp_sub"
      EndProperty
      BeginProperty Field80 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "cancelado"
         Caption         =   "cancelado"
      EndProperty
      BeginProperty Field81 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canc_data"
         Caption         =   "canc_data"
      EndProperty
      BeginProperty Field82 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "canc_usu"
         Caption         =   "canc_usu"
      EndProperty
      BeginProperty Field83 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "canc_obs"
         Caption         =   "canc_obs"
      EndProperty
      BeginProperty Field84 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_ctc_cif"
         Caption         =   "at_ctc_cif"
      EndProperty
      BeginProperty Field85 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_ctc_cif_data"
         Caption         =   "at_ctc_cif_data"
      EndProperty
      BeginProperty Field86 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "codevo"
         Caption         =   "codevo"
      EndProperty
      BeginProperty Field87 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_devol"
         Caption         =   "at_devol"
      EndProperty
      BeginProperty Field88 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_cliente"
         Caption         =   "at_cliente"
      EndProperty
      BeginProperty Field89 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_edi"
         Caption         =   "at_edi"
      EndProperty
      BeginProperty Field90 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_edi_data"
         Caption         =   "at_edi_data"
      EndProperty
      BeginProperty Field91 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_edisub"
         Caption         =   "at_edisub"
      EndProperty
      BeginProperty Field92 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_edisub_data"
         Caption         =   "at_edisub_data"
      EndProperty
      BeginProperty Field93 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "reimpr"
         Caption         =   "reimpr"
      EndProperty
      BeginProperty Field94 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "reimpr_data"
         Caption         =   "reimpr_data"
      EndProperty
      BeginProperty Field95 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "reimpr_usu"
         Caption         =   "reimpr_usu"
      EndProperty
      BeginProperty Field96 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field97 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "edi_nfsitla"
         Caption         =   "edi_nfsitla"
      EndProperty
      BeginProperty Field98 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "edi_nfsitladata"
         Caption         =   "edi_nfsitladata"
      EndProperty
      BeginProperty Field99 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ordem"
         Caption         =   "ordem"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset122 
      CommandName     =   "Sel_CTREmissor"
      CommDispId      =   2262
      RsDispId        =   2277
      CommandText     =   "select filialctc from tb_ctc_esp where tem_ocorr <> 'C' and emissor = ? and impresso = '' and tipodoc = 'MC'  order by ctc"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset123 
      CommandName     =   "Alt_ImpressoSimCTR"
      CommDispId      =   2268
      RsDispId        =   -1
      CommandText     =   "update tb_ctc_esp set impresso = 'S', impresso_data = getdate(), impresso_usu = ? where filialctc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "usuario"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset124 
      CommandName     =   "Ins_TempNFTb_Nf"
      CommDispId      =   2279
      RsDispId        =   -1
      CommandText     =   "insert into tb_tempnf (ctrnf, numnf, serie, valornf, pesonf, volumesnf, ordem) values ( ? , ? , ? , ? , ? , ? , ? )"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "ctr"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "numnf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "serie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "valor"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "peso"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "volumes"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         UserName        =   "ordem"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset125 
      CommandName     =   "Sel_NFsdoCTR"
      CommDispId      =   2281
      RsDispId        =   2286
      CommandText     =   "select * from tb_nf_esp  where filialctc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   21
      BeginProperty Field1 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "idcodigo"
         Caption         =   "idcodigo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field3 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numnfnum"
         Caption         =   "numnfnum"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "numnf"
         Caption         =   "numnf"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "serie"
         Caption         =   "serie"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cliente_cgc"
         Caption         =   "cliente_cgc"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cliente_nome"
         Caption         =   "cliente_nome"
      EndProperty
      BeginProperty Field8 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "emissao_nf"
         Caption         =   "emissao_nf"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "numpedido"
         Caption         =   "numpedido"
      EndProperty
      BeginProperty Field10 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtpedido"
         Caption         =   "dtpedido"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valornf"
         Caption         =   "valornf"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesonf"
         Caption         =   "pesonf"
      EndProperty
      BeginProperty Field13 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "volumesnf"
         Caption         =   "volumesnf"
      EndProperty
      BeginProperty Field14 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data_interface"
         Caption         =   "data_interface"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hora_interface"
         Caption         =   "hora_interface"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_cliente"
         Caption         =   "at_cliente"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "canhotonf"
         Caption         =   "canhotonf"
      EndProperty
      BeginProperty Field18 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "canhotonfprot"
         Caption         =   "canhotonfprot"
      EndProperty
      BeginProperty Field19 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canhotonfdata"
         Caption         =   "canhotonfdata"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tem_ocorrnf"
         Caption         =   "tem_ocorrnf"
      EndProperty
      BeginProperty Field21 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ordem"
         Caption         =   "ordem"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset126 
      CommandName     =   "Exc_NFsAlteracao"
      CommDispId      =   2287
      RsDispId        =   -1
      CommandText     =   "delete from tb_nf_esp where filialctc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset127 
      CommandName     =   "Alt_MinutaCTR"
      CommDispId      =   2289
      RsDispId        =   -1
      CommandText     =   "dbo.sp_alt_minutacompl"
      ActiveConnectionName=   "cn_informa"
      CallSyntax      =   $"de_informaEM.dsx":20FF
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   54
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@motivodoc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@prioridade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@prev_entrega"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@remet_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@remet_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@remet_end"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@remet_cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@remet_uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@remet_cep"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "@respons_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "@respons_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "@dest_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "@dest_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "@dest_end"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "@dest_cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "@dest_uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "@dest_cep"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "@cidade_orig"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "@uf_orig"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P22 
         RealName        =   "@via"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P23 
         RealName        =   "@cidade_dest"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P24 
         RealName        =   "@uf_dest"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P25 
         RealName        =   "@regiao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P26 
         RealName        =   "@regiaogeo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P27 
         RealName        =   "@regiaosac"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P28 
         RealName        =   "@atendsac"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P29 
         RealName        =   "@nfs"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   300
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P30 
         RealName        =   "@valmerc"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P31 
         RealName        =   "@peso"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P32 
         RealName        =   "@pesotax"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P33 
         RealName        =   "@volumes"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P34 
         RealName        =   "@especie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P35 
         RealName        =   "@natureza"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   25
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P36 
         RealName        =   "@naturezaobs"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   25
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P37 
         RealName        =   "@tabfrete"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P38 
         RealName        =   "@tabfretedescr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   100
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P39 
         RealName        =   "@fretepeso"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P40 
         RealName        =   "@fretevalor"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P41 
         RealName        =   "@gris"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P42 
         RealName        =   "@txcoleta"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P43 
         RealName        =   "@txentregared"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P44 
         RealName        =   "@txurgencia"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P45 
         RealName        =   "@pedagio"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P46 
         RealName        =   "@txoutros"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P47 
         RealName        =   "@descrtxoutros"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P48 
         RealName        =   "@fretetotal"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P49 
         RealName        =   "@fretetotalbruto"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P50 
         RealName        =   "@modal"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P51 
         RealName        =   "@obs_emissao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   320
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P52 
         RealName        =   "@fpag"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P53 
         RealName        =   "@redesp_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P54 
         RealName        =   "@redesp_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset128 
      CommandName     =   "Sel_Ajuste1"
      CommDispId      =   2325
      RsDispId        =   2328
      CommandText     =   $"de_informaEM.dsx":21C5
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset129 
      CommandName     =   "Exc_Manifesto"
      CommDispId      =   2336
      RsDispId        =   -1
      CommandText     =   "delete from tb_manifesto where filialmanifesto = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "filialmanifesto"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset130 
      CommandName     =   "Sel_CtrsParaReimprimir"
      CommDispId      =   2352
      RsDispId        =   2357
      CommandText     =   $"de_informaEM.dsx":22A5
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "remet_cgc"
         Caption         =   "remet_cgc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_nome"
         Caption         =   "remet_nome"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "qtde"
         Caption         =   "qtde"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset131 
      CommandName     =   "Sel_Ajustex"
      CommDispId      =   2366
      RsDispId        =   2383
      CommandText     =   "select * from plan1"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "CTC"
         Caption         =   "CTC"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "ACAO"
         Caption         =   "ACAO"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset132 
      CommandName     =   "Alt_AjusteD"
      CommDispId      =   2379
      RsDispId        =   -1
      CommandText     =   "update tb_nf_esp set  valornf = valornf / 1000 where filialctc = ? "
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset133 
      CommandName     =   "Alt_AjusteZ"
      CommDispId      =   2388
      RsDispId        =   -1
      CommandText     =   "update tb_nf_esp set  valornf = 0 where filialctc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset134 
      CommandName     =   "Sel_ImpFoxExcel"
      CommDispId      =   2390
      RsDispId        =   2397
      CommandText     =   "select * from tb_basecli where ordvenda = ? and item = ? and pedido = ? and numnf = ? and serie = ? "
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   13
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "ordvenda"
         Caption         =   "ordvenda"
      EndProperty
      BeginProperty Field2 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "item"
         Caption         =   "item"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "pedido"
         Caption         =   "pedido"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "codclinf"
         Caption         =   "codclinf"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "clientenf"
         Caption         =   "clientenf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidadenf"
         Caption         =   "cidadenf"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "ufnf"
         Caption         =   "ufnf"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "codmaterial"
         Caption         =   "codmaterial"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "material"
         Caption         =   "material"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "numnf"
         Caption         =   "numnf"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "serie"
         Caption         =   "serie"
      EndProperty
      BeginProperty Field12 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datanf"
         Caption         =   "datanf"
      EndProperty
      BeginProperty Field13 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dataarq"
         Caption         =   "dataarq"
      EndProperty
      NumGroups       =   0
      ParamCount      =   5
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "ordvenda"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "item"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "pedido"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "numnf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "serie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset135 
      CommandName     =   "Ins_ImpFoxExcel"
      CommDispId      =   2398
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":2550
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   13
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   80
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "Param11"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "Param12"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "Param13"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset136 
      CommandName     =   "Sel_NotFis"
      CommDispId      =   2400
      RsDispId        =   2405
      CommandText     =   "select * from tb_notfis where id_notfis = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   26
      BeginProperty Field1 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "id_notfis"
         Caption         =   "id_notfis"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "remet_cgc"
         Caption         =   "remet_cgc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_nome"
         Caption         =   "remet_nome"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "dest_cgc"
         Caption         =   "dest_cgc"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_nome"
         Caption         =   "dest_nome"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "dest_ie"
         Caption         =   "dest_ie"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_end"
         Caption         =   "dest_end"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "dest_bairro"
         Caption         =   "dest_bairro"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "dest_cidade"
         Caption         =   "dest_cidade"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "dest_uf"
         Caption         =   "dest_uf"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "dest_cep"
         Caption         =   "dest_cep"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "tipocarga"
         Caption         =   "tipocarga"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "tipofrete"
         Caption         =   "tipofrete"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "numnf"
         Caption         =   "numnf"
      EndProperty
      BeginProperty Field15 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numnfnum"
         Caption         =   "numnfnum"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "serie"
         Caption         =   "serie"
      EndProperty
      BeginProperty Field17 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "emissaonf"
         Caption         =   "emissaonf"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "natureza"
         Caption         =   "natureza"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "especie"
         Caption         =   "especie"
      EndProperty
      BeginProperty Field20 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "volumes"
         Caption         =   "volumes"
      EndProperty
      BeginProperty Field21 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valmerc"
         Caption         =   "valmerc"
      EndProperty
      BeginProperty Field22 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "peso"
         Caption         =   "peso"
      EndProperty
      BeginProperty Field23 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesocub"
         Caption         =   "pesocub"
      EndProperty
      BeginProperty Field24 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dataimp"
         Caption         =   "dataimp"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "emitido_auto"
         Caption         =   "emitido_auto"
      EndProperty
      BeginProperty Field26 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "emit_data"
         Caption         =   "emit_data"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "id"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset137 
      CommandName     =   "Ins_NotFis"
      CommDispId      =   2406
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":266B
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   27
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "Param11"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "Param12"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "Param13"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "Param14"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "Param15"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "Param16"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "Param17"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "Param18"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "Param19"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "Param20"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "Param21"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P22 
         RealName        =   "Param22"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P23 
         RealName        =   "Param23"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P24 
         RealName        =   "Param24"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P25 
         RealName        =   "Param25"
         Direction       =   1
         Precision       =   23
         Scale           =   3
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P26 
         RealName        =   "Param26"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P27 
         RealName        =   "Param27"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset138 
      CommandName     =   "Sel_AjustCorreios"
      CommDispId      =   2408
      RsDispId        =   2411
      CommandText     =   "SELECT * FROM CORREIO WHERE NUMNF NOT IN (SELECT NUMNF FROM TB_NF_ESP WHERE CLIENTE_CGC LIKE '04229761%') "
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "NUMNF"
         Caption         =   "NUMNF"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "DEST_CGC"
         Caption         =   "DEST_CGC"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "DEST_NOME"
         Caption         =   "DEST_NOME"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "CIDADE_DEST"
         Caption         =   "CIDADE_DEST"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "UF_DEST"
         Caption         =   "UF_DEST"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "END_DEST"
         Caption         =   "END_DEST"
      EndProperty
      BeginProperty Field7 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "VOLUMES"
         Caption         =   "VOLUMES"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "VALMERC"
         Caption         =   "VALMERC"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "PESO"
         Caption         =   "PESO"
      EndProperty
      BeginProperty Field10 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "QTDEFOX"
         Caption         =   "QTDEFOX"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset139 
      CommandName     =   "Sel_AjusteBuscaNF"
      CommDispId      =   2412
      RsDispId        =   2417
      CommandText     =   "select * from tb_nf_esp where numnf = ? and cliente_cgc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   21
      BeginProperty Field1 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "idcodigo"
         Caption         =   "idcodigo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field3 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numnfnum"
         Caption         =   "numnfnum"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "numnf"
         Caption         =   "numnf"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "serie"
         Caption         =   "serie"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cliente_cgc"
         Caption         =   "cliente_cgc"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cliente_nome"
         Caption         =   "cliente_nome"
      EndProperty
      BeginProperty Field8 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "emissao_nf"
         Caption         =   "emissao_nf"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "numpedido"
         Caption         =   "numpedido"
      EndProperty
      BeginProperty Field10 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtpedido"
         Caption         =   "dtpedido"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valornf"
         Caption         =   "valornf"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesonf"
         Caption         =   "pesonf"
      EndProperty
      BeginProperty Field13 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "volumesnf"
         Caption         =   "volumesnf"
      EndProperty
      BeginProperty Field14 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data_interface"
         Caption         =   "data_interface"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hora_interface"
         Caption         =   "hora_interface"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_cliente"
         Caption         =   "at_cliente"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "canhotonf"
         Caption         =   "canhotonf"
      EndProperty
      BeginProperty Field18 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "canhotonfprot"
         Caption         =   "canhotonfprot"
      EndProperty
      BeginProperty Field19 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canhotonfdata"
         Caption         =   "canhotonfdata"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tem_ocorrnf"
         Caption         =   "tem_ocorrnf"
      EndProperty
      BeginProperty Field21 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ordem"
         Caption         =   "ordem"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "NF"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   12
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "Remet"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset140 
      CommandName     =   "Alt_ImpFoxExcel"
      CommDispId      =   2418
      RsDispId        =   -1
      CommandText     =   "update tb_basecli set entr_solic = ?, acaofox = ?, pacote = ? where numnf = ? "
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   4
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   200
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset141 
      CommandName     =   "Sel_NFNotFis"
      CommDispId      =   2420
      RsDispId        =   2426
      CommandText     =   "select * from tb_notfis where remet_cgc = ? and numnfnum = ? and serie = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   27
      BeginProperty Field1 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "id_notfis"
         Caption         =   "id_notfis"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "remet_cgc"
         Caption         =   "remet_cgc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_nome"
         Caption         =   "remet_nome"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "dest_cgc"
         Caption         =   "dest_cgc"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_nome"
         Caption         =   "dest_nome"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "dest_ie"
         Caption         =   "dest_ie"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_end"
         Caption         =   "dest_end"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "dest_bairro"
         Caption         =   "dest_bairro"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "dest_cidade"
         Caption         =   "dest_cidade"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "dest_uf"
         Caption         =   "dest_uf"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "dest_cep"
         Caption         =   "dest_cep"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "tipocarga"
         Caption         =   "tipocarga"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "tipofrete"
         Caption         =   "tipofrete"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "numnf"
         Caption         =   "numnf"
      EndProperty
      BeginProperty Field15 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numnfnum"
         Caption         =   "numnfnum"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "serie"
         Caption         =   "serie"
      EndProperty
      BeginProperty Field17 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "emissaonf"
         Caption         =   "emissaonf"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "natureza"
         Caption         =   "natureza"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "especie"
         Caption         =   "especie"
      EndProperty
      BeginProperty Field20 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "volumes"
         Caption         =   "volumes"
      EndProperty
      BeginProperty Field21 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valmerc"
         Caption         =   "valmerc"
      EndProperty
      BeginProperty Field22 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "peso"
         Caption         =   "peso"
      EndProperty
      BeginProperty Field23 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesocub"
         Caption         =   "pesocub"
      EndProperty
      BeginProperty Field24 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datainterface"
         Caption         =   "datainterface"
      EndProperty
      BeginProperty Field25 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dataimp"
         Caption         =   "dataimp"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "emitido_auto"
         Caption         =   "emitido_auto"
      EndProperty
      BeginProperty Field27 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "emit_data"
         Caption         =   "emit_data"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "nfnum"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "serie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset142 
      CommandName     =   "Alt_AjusteQtde"
      CommDispId      =   2427
      RsDispId        =   -1
      CommandText     =   "update tb_notfis set qtdeitem = ? where numnfnum = ? and serie = ? and id_notfis = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   4
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset143 
      CommandName     =   "Alt_AcertoNotFis"
      CommDispId      =   2435
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":280E
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   24
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "id"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "cgccli"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "nomecli"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "ie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "end"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "bairro"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         UserName        =   "cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         UserName        =   "uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         UserName        =   "cep"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         UserName        =   "tipocarga"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "Param11"
         UserName        =   "tipofrete"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "Param12"
         UserName        =   "emissaonf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "Param13"
         UserName        =   "natureza"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "Param14"
         UserName        =   "especie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "Param15"
         UserName        =   "volumes"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "Param16"
         UserName        =   "valmerc"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "Param17"
         UserName        =   "peso"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "Param18"
         UserName        =   "pesocub"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "Param19"
         UserName        =   "dataimp"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "Param20"
         UserName        =   "datainterface"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "Param21"
         UserName        =   "emitido"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P22 
         RealName        =   "Param22"
         UserName        =   "qtde"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P23 
         RealName        =   "Param23"
         UserName        =   "serie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P24 
         RealName        =   "Param24"
         UserName        =   "nfnum"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset144 
      CommandName     =   "Sel_NotFisItem"
      CommDispId      =   2443
      RsDispId        =   2450
      CommandText     =   "select * from tb_notfisitem where numnfnum = ? and serie = ? and posicao = ? "
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "id_notfis"
         Caption         =   "id_notfis"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "remet_cgc"
         Caption         =   "remet_cgc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_nome"
         Caption         =   "remet_nome"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "numnf"
         Caption         =   "numnf"
      EndProperty
      BeginProperty Field5 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numnfnum"
         Caption         =   "numnfnum"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "serie"
         Caption         =   "serie"
      EndProperty
      BeginProperty Field7 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "posicao"
         Caption         =   "posicao"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "codigoitem"
         Caption         =   "codigoitem"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "descricaoitem"
         Caption         =   "descricaoitem"
      EndProperty
      BeginProperty Field10 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "qtdeitem"
         Caption         =   "qtdeitem"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "NFNum"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "Serie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "Posicao"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset145 
      CommandName     =   "Ins_NotFisItem"
      CommDispId      =   2451
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":2984
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   10
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "id"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "remet_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "remet_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "numnf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "numnfnum"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "serie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         UserName        =   "Posicao"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         UserName        =   "CodItem"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         UserName        =   "DescrItem"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         UserName        =   "qtde"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset146 
      CommandName     =   "Alt_NotFisItem"
      CommDispId      =   2465
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":2A3B
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "id"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "coditem"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "descricao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "qtde"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "numnfnum"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "serie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         UserName        =   "posicao"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset147 
      CommandName     =   "Sel_BuscaNFSerieCGC"
      CommDispId      =   2474
      RsDispId        =   2479
      CommandText     =   $"de_informaEM.dsx":2AC7
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   138
      BeginProperty Field1 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "tipodoc"
         Caption         =   "tipodoc"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "motivodoc"
         Caption         =   "motivodoc"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "filial"
         Caption         =   "filial"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "ctc"
         Caption         =   "ctc"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hora"
         Caption         =   "hora"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "emissor"
         Caption         =   "emissor"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "conferente"
         Caption         =   "conferente"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "prioridade"
         Caption         =   "prioridade"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "prev_entrega"
         Caption         =   "prev_entrega"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "remet_cgc"
         Caption         =   "remet_cgc"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_nome"
         Caption         =   "remet_nome"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "remet_end"
         Caption         =   "remet_end"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "remet_cidade"
         Caption         =   "remet_cidade"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "remet_uf"
         Caption         =   "remet_uf"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "remet_cep"
         Caption         =   "remet_cep"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "respons_cgc"
         Caption         =   "respons_cgc"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "respons_nome"
         Caption         =   "respons_nome"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "dest_cgc"
         Caption         =   "dest_cgc"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_nome"
         Caption         =   "dest_nome"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "dest_end"
         Caption         =   "dest_end"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "dest_cidade"
         Caption         =   "dest_cidade"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "dest_uf"
         Caption         =   "dest_uf"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "dest_cep"
         Caption         =   "dest_cep"
      EndProperty
      BeginProperty Field26 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade_orig"
         Caption         =   "cidade_orig"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_orig"
         Caption         =   "uf_orig"
      EndProperty
      BeginProperty Field28 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "via"
         Caption         =   "via"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade_dest"
         Caption         =   "cidade_dest"
      EndProperty
      BeginProperty Field30 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_dest"
         Caption         =   "uf_dest"
      EndProperty
      BeginProperty Field31 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiao"
         Caption         =   "regiao"
      EndProperty
      BeginProperty Field32 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      BeginProperty Field33 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "regiaosac"
         Caption         =   "regiaosac"
      EndProperty
      BeginProperty Field34 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "atendsac"
         Caption         =   "atendsac"
      EndProperty
      BeginProperty Field35 
         Precision       =   0
         Size            =   300
         Scale           =   0
         Type            =   200
         Name            =   "nfs"
         Caption         =   "nfs"
      EndProperty
      BeginProperty Field36 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valmerc"
         Caption         =   "valmerc"
      EndProperty
      BeginProperty Field37 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "peso"
         Caption         =   "peso"
      EndProperty
      BeginProperty Field38 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesotax"
         Caption         =   "pesotax"
      EndProperty
      BeginProperty Field39 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "volumes"
         Caption         =   "volumes"
      EndProperty
      BeginProperty Field40 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "especie"
         Caption         =   "especie"
      EndProperty
      BeginProperty Field41 
         Precision       =   0
         Size            =   25
         Scale           =   0
         Type            =   200
         Name            =   "natureza"
         Caption         =   "natureza"
      EndProperty
      BeginProperty Field42 
         Precision       =   0
         Size            =   25
         Scale           =   0
         Type            =   200
         Name            =   "naturezaobs"
         Caption         =   "naturezaobs"
      EndProperty
      BeginProperty Field43 
         Precision       =   0
         Size            =   29
         Scale           =   0
         Type            =   200
         Name            =   "dimensoes"
         Caption         =   "dimensoes"
      EndProperty
      BeginProperty Field44 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "tabfrete"
         Caption         =   "tabfrete"
      EndProperty
      BeginProperty Field45 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "tabfretedescr"
         Caption         =   "tabfretedescr"
      EndProperty
      BeginProperty Field46 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepeso"
         Caption         =   "fretepeso"
      EndProperty
      BeginProperty Field47 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalor"
         Caption         =   "fretevalor"
      EndProperty
      BeginProperty Field48 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gris"
         Caption         =   "gris"
      EndProperty
      BeginProperty Field49 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txentregared"
         Caption         =   "txentregared"
      EndProperty
      BeginProperty Field50 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txurgencia"
         Caption         =   "txurgencia"
      EndProperty
      BeginProperty Field51 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pedagio"
         Caption         =   "pedagio"
      EndProperty
      BeginProperty Field52 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretenacional"
         Caption         =   "fretenacional"
      EndProperty
      BeginProperty Field53 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "advalorem"
         Caption         =   "advalorem"
      EndProperty
      BeginProperty Field54 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txorigem"
         Caption         =   "txorigem"
      EndProperty
      BeginProperty Field55 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txdestino"
         Caption         =   "txdestino"
      EndProperty
      BeginProperty Field56 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txredespacho"
         Caption         =   "txredespacho"
      EndProperty
      BeginProperty Field57 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txcoleta"
         Caption         =   "txcoleta"
      EndProperty
      BeginProperty Field58 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txoutros"
         Caption         =   "txoutros"
      EndProperty
      BeginProperty Field59 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "descrtxoutros"
         Caption         =   "descrtxoutros"
      EndProperty
      BeginProperty Field60 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotal"
         Caption         =   "fretetotal"
      EndProperty
      BeginProperty Field61 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotalbruto"
         Caption         =   "fretetotalbruto"
      EndProperty
      BeginProperty Field62 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepagosub"
         Caption         =   "fretepagosub"
      EndProperty
      BeginProperty Field63 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tribut"
         Caption         =   "tribut"
      EndProperty
      BeginProperty Field64 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "aliquota"
         Caption         =   "aliquota"
      EndProperty
      BeginProperty Field65 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "filialmanifesto"
         Caption         =   "filialmanifesto"
      EndProperty
      BeginProperty Field66 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "fpag"
         Caption         =   "fpag"
      EndProperty
      BeginProperty Field67 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "faturar"
         Caption         =   "faturar"
      EndProperty
      BeginProperty Field68 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "faturanum"
         Caption         =   "faturanum"
      EndProperty
      BeginProperty Field69 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "entrega_data"
         Caption         =   "entrega_data"
      EndProperty
      BeginProperty Field70 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "recebedor"
         Caption         =   "recebedor"
      EndProperty
      BeginProperty Field71 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   129
         Name            =   "entrega_hora"
         Caption         =   "entrega_hora"
      EndProperty
      BeginProperty Field72 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "modal"
         Caption         =   "modal"
      EndProperty
      BeginProperty Field73 
         Precision       =   0
         Size            =   320
         Scale           =   0
         Type            =   200
         Name            =   "obs_emissao"
         Caption         =   "obs_emissao"
      EndProperty
      BeginProperty Field74 
         Precision       =   0
         Size            =   186
         Scale           =   0
         Type            =   200
         Name            =   "obs_ocorr"
         Caption         =   "obs_ocorr"
      EndProperty
      BeginProperty Field75 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "arq"
         Caption         =   "arq"
      EndProperty
      BeginProperty Field76 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tem_ocorr"
         Caption         =   "tem_ocorr"
      EndProperty
      BeginProperty Field77 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "redesp_cgc"
         Caption         =   "redesp_cgc"
      EndProperty
      BeginProperty Field78 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "redesp_nome"
         Caption         =   "redesp_nome"
      EndProperty
      BeginProperty Field79 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "transp_sub"
         Caption         =   "transp_sub"
      EndProperty
      BeginProperty Field80 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "cancelado"
         Caption         =   "cancelado"
      EndProperty
      BeginProperty Field81 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canc_data"
         Caption         =   "canc_data"
      EndProperty
      BeginProperty Field82 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "canc_usu"
         Caption         =   "canc_usu"
      EndProperty
      BeginProperty Field83 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "canc_obs"
         Caption         =   "canc_obs"
      EndProperty
      BeginProperty Field84 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_ctc_cif"
         Caption         =   "at_ctc_cif"
      EndProperty
      BeginProperty Field85 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_ctc_cif_data"
         Caption         =   "at_ctc_cif_data"
      EndProperty
      BeginProperty Field86 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "codevo"
         Caption         =   "codevo"
      EndProperty
      BeginProperty Field87 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_devol"
         Caption         =   "at_devol"
      EndProperty
      BeginProperty Field88 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_cliente"
         Caption         =   "at_cliente"
      EndProperty
      BeginProperty Field89 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_edi"
         Caption         =   "at_edi"
      EndProperty
      BeginProperty Field90 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_edi_data"
         Caption         =   "at_edi_data"
      EndProperty
      BeginProperty Field91 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_edisub"
         Caption         =   "at_edisub"
      EndProperty
      BeginProperty Field92 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "at_edisub_data"
         Caption         =   "at_edisub_data"
      EndProperty
      BeginProperty Field93 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "impresso"
         Caption         =   "impresso"
      EndProperty
      BeginProperty Field94 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "impresso_data"
         Caption         =   "impresso_data"
      EndProperty
      BeginProperty Field95 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "impresso_usu"
         Caption         =   "impresso_usu"
      EndProperty
      BeginProperty Field96 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "reimpr"
         Caption         =   "reimpr"
      EndProperty
      BeginProperty Field97 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "reimpr_data"
         Caption         =   "reimpr_data"
      EndProperty
      BeginProperty Field98 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "reimpr_usu"
         Caption         =   "reimpr_usu"
      EndProperty
      BeginProperty Field99 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field100 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "edi_nfsitla"
         Caption         =   "edi_nfsitla"
      EndProperty
      BeginProperty Field101 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "edi_nfsitladata"
         Caption         =   "edi_nfsitladata"
      EndProperty
      BeginProperty Field102 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ordem"
         Caption         =   "ordem"
      EndProperty
      BeginProperty Field103 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valmercCTR03"
         Caption         =   "valmercCTR03"
      EndProperty
      BeginProperty Field104 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesoCTR03"
         Caption         =   "pesoCTR03"
      EndProperty
      BeginProperty Field105 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesotaxCTR03"
         Caption         =   "pesotaxCTR03"
      EndProperty
      BeginProperty Field106 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "volumesCTR03"
         Caption         =   "volumesCTR03"
      EndProperty
      BeginProperty Field107 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepesoCTR03"
         Caption         =   "fretepesoCTR03"
      EndProperty
      BeginProperty Field108 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalorCTR03"
         Caption         =   "fretevalorCTR03"
      EndProperty
      BeginProperty Field109 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "grisCTR03"
         Caption         =   "grisCTR03"
      EndProperty
      BeginProperty Field110 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txentregaredCTR03"
         Caption         =   "txentregaredCTR03"
      EndProperty
      BeginProperty Field111 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txurgenciaCTR03"
         Caption         =   "txurgenciaCTR03"
      EndProperty
      BeginProperty Field112 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pedagioCTR03"
         Caption         =   "pedagioCTR03"
      EndProperty
      BeginProperty Field113 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txcoletaCTR03"
         Caption         =   "txcoletaCTR03"
      EndProperty
      BeginProperty Field114 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "txoutrosCTR03"
         Caption         =   "txoutrosCTR03"
      EndProperty
      BeginProperty Field115 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "descrtxoutrosCTR03"
         Caption         =   "descrtxoutrosCTR03"
      EndProperty
      BeginProperty Field116 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotalCTR03"
         Caption         =   "fretetotalCTR03"
      EndProperty
      BeginProperty Field117 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretetotalbrutoCTR03"
         Caption         =   "fretetotalbrutoCTR03"
      EndProperty
      BeginProperty Field118 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "idcodigo"
         Caption         =   "idcodigo"
      EndProperty
      BeginProperty Field119 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "filialctc"
         Caption         =   "filialctc"
      EndProperty
      BeginProperty Field120 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numnfnum"
         Caption         =   "numnfnum"
      EndProperty
      BeginProperty Field121 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "numnf"
         Caption         =   "numnf"
      EndProperty
      BeginProperty Field122 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "serie"
         Caption         =   "serie"
      EndProperty
      BeginProperty Field123 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "cliente_cgc"
         Caption         =   "cliente_cgc"
      EndProperty
      BeginProperty Field124 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "cliente_nome"
         Caption         =   "cliente_nome"
      EndProperty
      BeginProperty Field125 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "emissao_nf"
         Caption         =   "emissao_nf"
      EndProperty
      BeginProperty Field126 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "numpedido"
         Caption         =   "numpedido"
      EndProperty
      BeginProperty Field127 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtpedido"
         Caption         =   "dtpedido"
      EndProperty
      BeginProperty Field128 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "valornf"
         Caption         =   "valornf"
      EndProperty
      BeginProperty Field129 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesonf"
         Caption         =   "pesonf"
      EndProperty
      BeginProperty Field130 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "volumesnf"
         Caption         =   "volumesnf"
      EndProperty
      BeginProperty Field131 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data_interface"
         Caption         =   "data_interface"
      EndProperty
      BeginProperty Field132 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "hora_interface"
         Caption         =   "hora_interface"
      EndProperty
      BeginProperty Field133 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "at_cliente"
         Caption         =   "at_cliente"
      EndProperty
      BeginProperty Field134 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "canhotonf"
         Caption         =   "canhotonf"
      EndProperty
      BeginProperty Field135 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "canhotonfprot"
         Caption         =   "canhotonfprot"
      EndProperty
      BeginProperty Field136 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "canhotonfdata"
         Caption         =   "canhotonfdata"
      EndProperty
      BeginProperty Field137 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "tem_ocorrnf"
         Caption         =   "tem_ocorrnf"
      EndProperty
      BeginProperty Field138 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "ordem"
         Caption         =   "ordem"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "nfnum"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   19
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "serie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset148 
      CommandName     =   "Sel_ClassificaCFOPTodos"
      CommDispId      =   2480
      RsDispId        =   2485
      CommandText     =   "select distinct classifica from tb_cfop where status = '1'"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "classifica"
         Caption         =   "classifica"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset149 
      CommandName     =   "Sel_UFsFiscal"
      CommDispId      =   2486
      RsDispId        =   2491
      CommandText     =   "select * from tb_caduffisco where uf_orig = ? and uf_dest = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   6
      BeginProperty Field1 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_orig"
         Caption         =   "uf_orig"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "rgeo_orig"
         Caption         =   "rgeo_orig"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf_dest"
         Caption         =   "uf_dest"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "rgeo_dest"
         Caption         =   "rgeo_dest"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "aliquota"
         Caption         =   "aliquota"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "subtribut"
         Caption         =   "subtribut"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "ufOrig"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "ufDest"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset150 
      CommandName     =   "Ins_Ctc_Ctr"
      CommDispId      =   2494
      RsDispId        =   -1
      CommandText     =   "dbo.sp_ins_ctc_ctr"
      ActiveConnectionName=   "cn_informa"
      CallSyntax      =   $"de_informaEM.dsx":2B7C
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   58
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@tipodoc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@motivodoc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@filial"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@ctc"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@data"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@hora"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@prioridade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@controleRU"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@prev_entrega"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "@prev_entregatipo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "@remet_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "@remet_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "@remet_end"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "@remet_cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "@remet_uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "@remet_cep"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "@remet_ie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "@respons_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "@respons_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P22 
         RealName        =   "@respons_end"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P23 
         RealName        =   "@respons_cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P24 
         RealName        =   "@respons_uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P25 
         RealName        =   "@respons_ie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P26 
         RealName        =   "@dest_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P27 
         RealName        =   "@dest_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P28 
         RealName        =   "@dest_end"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P29 
         RealName        =   "@dest_cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P30 
         RealName        =   "@dest_uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P31 
         RealName        =   "@dest_cep"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P32 
         RealName        =   "@dest_ie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P33 
         RealName        =   "@cidade_orig"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P34 
         RealName        =   "@uf_orig"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P35 
         RealName        =   "@via"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P36 
         RealName        =   "@cidade_dest"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P37 
         RealName        =   "@uf_dest"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P38 
         RealName        =   "@regiao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P39 
         RealName        =   "@regiaogeo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P40 
         RealName        =   "@regiaosac"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P41 
         RealName        =   "@atendsac"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P42 
         RealName        =   "@nfs"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   300
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P43 
         RealName        =   "@valmerc"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P44 
         RealName        =   "@peso"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P45 
         RealName        =   "@pesotax"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P46 
         RealName        =   "@volumes"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P47 
         RealName        =   "@especie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P48 
         RealName        =   "@natureza"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   25
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P49 
         RealName        =   "@naturezaobs"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   25
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P50 
         RealName        =   "@modal"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P51 
         RealName        =   "@obs_emissao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   320
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P52 
         RealName        =   "@fpag"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P53 
         RealName        =   "@emissor"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P54 
         RealName        =   "@conferente"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P55 
         RealName        =   "@status"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P56 
         RealName        =   "@redesp_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P57 
         RealName        =   "@redesp_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P58 
         RealName        =   "@nservico"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset151 
      CommandName     =   "Ins_Ctc_Ctr2"
      CommDispId      =   2500
      RsDispId        =   -1
      CommandText     =   "dbo.sp_ins_ctc_ctr2"
      ActiveConnectionName=   "cn_informa"
      CallSyntax      =   "{? = CALL dbo.sp_ins_ctc_ctr2( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   28
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@fretepeso"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@fretepesobr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@fretevalor"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@fretevalorbr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@gris"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@grisbr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@txcoleta"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@txcoletabr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@txentregared"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "@txentregaredbr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "@txurgencia"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "@txurgenciabr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "@pedagio"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "@pedagiobr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "@txoutros"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "@txoutrosbr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "@descrtxoutros"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "@fretetotal"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "@fretetotalbruto"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P22 
         RealName        =   "@tabfrete"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P23 
         RealName        =   "@tabfretedescr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   100
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P24 
         RealName        =   "@cfop"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P25 
         RealName        =   "@tribut"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P26 
         RealName        =   "@aliquota"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P27 
         RealName        =   "@subtrib"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P28 
         RealName        =   "@subcontratacao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset152 
      CommandName     =   "Alt_Ctr"
      CommDispId      =   2502
      RsDispId        =   -1
      CommandText     =   "dbo.sp_alt_ctr"
      ActiveConnectionName=   "cn_informa"
      CallSyntax      =   $"de_informaEM.dsx":2C4A
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   57
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@motivodoc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@prioridade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@controleRU"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@prev_entrega"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@prev_entregatipo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@remet_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@remet_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@remet_end"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@remet_cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "@remet_uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "@remet_cep"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "@remet_ie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "@respons_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "@respons_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "@respons_end"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "@respons_cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "@respons_uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "@respons_ie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "@dest_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P22 
         RealName        =   "@dest_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P23 
         RealName        =   "@dest_end"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P24 
         RealName        =   "@dest_cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P25 
         RealName        =   "@dest_uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P26 
         RealName        =   "@dest_cep"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P27 
         RealName        =   "@dest_ie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P28 
         RealName        =   "@cidade_orig"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P29 
         RealName        =   "@uf_orig"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P30 
         RealName        =   "@via"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P31 
         RealName        =   "@cidade_dest"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P32 
         RealName        =   "@uf_dest"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P33 
         RealName        =   "@regiao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P34 
         RealName        =   "@regiaogeo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P35 
         RealName        =   "@regiaosac"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P36 
         RealName        =   "@atendsac"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P37 
         RealName        =   "@nfs"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   300
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P38 
         RealName        =   "@valmerc"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P39 
         RealName        =   "@peso"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P40 
         RealName        =   "@pesotax"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P41 
         RealName        =   "@volumes"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P42 
         RealName        =   "@especie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P43 
         RealName        =   "@natureza"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   25
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P44 
         RealName        =   "@naturezaobs"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   25
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P45 
         RealName        =   "@tabfrete"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P46 
         RealName        =   "@tabfretedescr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   100
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P47 
         RealName        =   "@cfop"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P48 
         RealName        =   "@tribut"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P49 
         RealName        =   "@aliquota"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P50 
         RealName        =   "@subtrib"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P51 
         RealName        =   "@subcontratacao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P52 
         RealName        =   "@modal"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P53 
         RealName        =   "@obs_emissao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   320
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P54 
         RealName        =   "@fpag"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P55 
         RealName        =   "@redesp_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P56 
         RealName        =   "@redesp_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P57 
         RealName        =   "@conferente"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset153 
      CommandName     =   "Alt_Ctr2"
      CommDispId      =   2507
      RsDispId        =   -1
      CommandText     =   "dbo.sp_alt_ctr2"
      ActiveConnectionName=   "cn_informa"
      CallSyntax      =   "{? = CALL dbo.sp_alt_ctr2( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   21
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@fretepeso"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@fretepesobr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@fretevalor"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@fretevalorbr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@gris"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@grisbr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@txcoleta"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@txcoletabr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@txentregared"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "@txentregaredbr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "@txurgencia"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "@txurgenciabr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "@pedagio"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "@pedagiobr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "@txoutros"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "@txoutrosbr"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "@descrtxoutros"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "@fretetotal"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "@fretetotalbruto"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset154 
      CommandName     =   "Alt_Ctc"
      CommDispId      =   2508
      RsDispId        =   -1
      CommandText     =   "dbo.sp_alt_ctc"
      ActiveConnectionName=   "cn_informa"
      CallSyntax      =   "{? = CALL dbo.sp_alt_ctc( ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   5
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@obs_emissao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   320
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@redesp_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@redesp_nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset155 
      CommandName     =   "Ins_Ocorr4Cod00"
      CommDispId      =   2510
      RsDispId        =   -1
      CommandText     =   "dbo.sp_ins_ocorr4cod00"
      ActiveConnectionName=   "cn_informa"
      CallSyntax      =   "{? = CALL dbo.sp_ins_ocorr4cod00( ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   10
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@emissaoctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@remet_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@cod_ocorr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@descr_ocorr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@data"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@hora"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@usu_ocorr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@usu_dataocorr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset156 
      CommandName     =   "Ins_Ocorr4"
      CommDispId      =   2512
      RsDispId        =   -1
      CommandText     =   "dbo.sp_ins_ocorr4"
      ActiveConnectionName=   "cn_informa"
      CallSyntax      =   "{? = CALL dbo.sp_ins_ocorr4( ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   10
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@emissaoctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@remet_cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@cod_ocorr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@descr_ocorr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@data"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@hora"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@usu_ocorr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@usu_dataocorr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset157 
      CommandName     =   "Ins_CadClientes"
      CommDispId      =   2514
      RsDispId        =   -1
      CommandText     =   "dbo.sp_ins_cadclientes"
      ActiveConnectionName=   "cn_informa"
      CallSyntax      =   $"de_informaEM.dsx":2D11
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   43
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@fantasia"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@apelido"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@endereco"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@complemento"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@cep"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@ie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "@pabx"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "@fax"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "@contato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "@fonecontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "@emailcontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "@anivercontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "@avisarcontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "@avusucontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "@contato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "@fonecontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P22 
         RealName        =   "@emailcontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P23 
         RealName        =   "@anivercontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P24 
         RealName        =   "@avisarcontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P25 
         RealName        =   "@avusucontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P26 
         RealName        =   "@contato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P27 
         RealName        =   "@fonecontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P28 
         RealName        =   "@emailcontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P29 
         RealName        =   "@anivercontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P30 
         RealName        =   "@avisarcontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P31 
         RealName        =   "@avusucontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P32 
         RealName        =   "@consigentrega"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P33 
         RealName        =   "@consigentregaair"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P34 
         RealName        =   "@consigtransf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P35 
         RealName        =   "@consigdevol"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P36 
         RealName        =   "@atendusu"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P37 
         RealName        =   "@prazo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   6
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P38 
         RealName        =   "@usuariocad"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P39 
         RealName        =   "@rem_des_log"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P40 
         RealName        =   "@alarm_ger"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P41 
         RealName        =   "@cfop"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P42 
         RealName        =   "@classefiscal"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P43 
         RealName        =   "@pessoafj"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset158 
      CommandName     =   "Sel_ClassificaCFOP"
      CommDispId      =   2516
      RsDispId        =   2521
      CommandText     =   "select * from tb_cfop where classifica = ? and uf = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   18
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "ID"
         Caption         =   "ID"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   4
         Scale           =   0
         Type            =   200
         Name            =   "cfop"
         Caption         =   "cfop"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "classifica"
         Caption         =   "classifica"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "classifica"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset159 
      CommandName     =   "Alt_CadCliente"
      CommDispId      =   2522
      RsDispId        =   -1
      CommandText     =   "dbo.sp_alt_cadcliente"
      ActiveConnectionName=   "cn_informa"
      CallSyntax      =   $"de_informaEM.dsx":2DB6
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   42
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@nome"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@fantasia"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@apelido"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@endereco"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   40
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@complemento"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@cep"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   8
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@ie"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "@pabx"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "@fax"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "@contato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "@fonecontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "@emailcontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "@anivercontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "@avisarcontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "@avusucontato1"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "@contato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "@fonecontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P22 
         RealName        =   "@emailcontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P23 
         RealName        =   "@anivercontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P24 
         RealName        =   "@avisarcontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P25 
         RealName        =   "@avusucontato2"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P26 
         RealName        =   "@contato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P27 
         RealName        =   "@fonecontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   15
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P28 
         RealName        =   "@emailcontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   30
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P29 
         RealName        =   "@anivercontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   5
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P30 
         RealName        =   "@avisarcontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P31 
         RealName        =   "@avusucontato3"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P32 
         RealName        =   "@consigentrega"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P33 
         RealName        =   "@consigentregaair"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P34 
         RealName        =   "@consigtransf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P35 
         RealName        =   "@consigdevol"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P36 
         RealName        =   "@atendusu"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P37 
         RealName        =   "@prazo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   6
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P38 
         RealName        =   "@rem_des_log"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P39 
         RealName        =   "@alarm_ger"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P40 
         RealName        =   "@cfop"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P41 
         RealName        =   "@classefiscal"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P42 
         RealName        =   "@pessoafj"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset160 
      CommandName     =   "Alt_TemOcorr_SN"
      CommDispId      =   2524
      RsDispId        =   -1
      CommandText     =   "dbo.sp_alt_temocorr_sn"
      ActiveConnectionName=   "cn_informa"
      CallSyntax      =   "{? = CALL dbo.sp_alt_temocorr_sn( ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@tem_ocorr"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@filialctc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset161 
      CommandName     =   "Alt_CadCliPessoaClasse"
      CommDispId      =   2526
      RsDispId        =   -1
      CommandText     =   "update tb_cadcli set pessoafj = ?, cfop = ?, classefiscal = ? where cgc = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   4
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "pessoa"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "cfop"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "classe"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "cgc"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   14
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset162 
      CommandName     =   "Sel_tempTR02"
      CommDispId      =   2528
      RsDispId        =   2565
      CommandText     =   "select * from tb_temptr02 where codigo = ? order by data "
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   20
      BeginProperty Field1 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "cim"
         Caption         =   "cim"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesode"
         Caption         =   "pesode"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesoate"
         Caption         =   "pesoate"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepeso"
         Caption         =   "fretepeso"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "porkilo"
         Caption         =   "porkilo"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalor"
         Caption         =   "fretevalor"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaate"
         Caption         =   "gen_txcoletaate"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletavalor"
         Caption         =   "gen_txcoletavalor"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaexced"
         Caption         =   "gen_txcoletaexced"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaate"
         Caption         =   "gen_txentregaate"
      EndProperty
      BeginProperty Field18 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregavalor"
         Caption         =   "gen_txentregavalor"
      EndProperty
      BeginProperty Field19 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaexced"
         Caption         =   "gen_txentregaexced"
      EndProperty
      BeginProperty Field20 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset163 
      CommandName     =   "Ins_TempTR02"
      CommDispId      =   2535
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":2E57
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   19
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   100
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "Param11"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "Param12"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "Param13"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "Param14"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "Param15"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "Param16"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "Param17"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "Param18"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "Param19"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset164 
      CommandName     =   "Sel_TempTR02Confere"
      CommDispId      =   2553
      RsDispId        =   2558
      CommandText     =   "select * from tb_temptr01 where codigo = ? and uf = ? and cim = ? and cidade = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "cim"
         Caption         =   "cim"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretemin"
         Caption         =   "fretemin"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "tarifaperc"
         Caption         =   "tarifaperc"
      EndProperty
      BeginProperty Field10 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "data"
         Caption         =   "data"
      EndProperty
      NumGroups       =   0
      ParamCount      =   4
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "CIM"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "Cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset165 
      CommandName     =   "Sel_UFnaoTratTR02"
      CommDispId      =   2559
      RsDispId        =   2563
      CommandText     =   "dbo.sp_sel_ufsnaotratTR02"
      ActiveConnectionName=   "cn_informa"
      CallSyntax      =   "{? = CALL dbo.sp_sel_ufsnaotratTR02( ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "nome"
         Caption         =   "nome"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "regiaogeo"
         Caption         =   "regiaogeo"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "@RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset166 
      CommandName     =   "Exc_TempTR02PesoDeAte"
      CommDispId      =   2567
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":2FC0
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   6
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "cod"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "pesode"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "pesoate"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "cim"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset167 
      CommandName     =   "Exc_TempTR02Tudo"
      CommDispId      =   2569
      RsDispId        =   -1
      CommandText     =   "delete from tb_temptr02 where codigo = ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset168 
      CommandName     =   "Ins_TR02Oficial"
      CommDispId      =   2571
      RsDispId        =   -1
      CommandText     =   $"de_informaEM.dsx":3034
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   22
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "origem"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "origemuf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "descricao"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   100
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "Param6"
         UserName        =   "cim"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "Param7"
         UserName        =   "cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "Param8"
         UserName        =   "pesode"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "Param9"
         UserName        =   "pesoate"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "Param10"
         UserName        =   "fretepeso"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "Param11"
         UserName        =   "porkilo"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "Param12"
         UserName        =   "complemento"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "Param13"
         UserName        =   "fretevalor"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "Param14"
         UserName        =   "coletaate"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "Param15"
         UserName        =   "coletavalor"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P16 
         RealName        =   "Param16"
         UserName        =   "coletaexced"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P17 
         RealName        =   "Param17"
         UserName        =   "entregaate"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P18 
         RealName        =   "Param18"
         UserName        =   "entregavalor"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P19 
         RealName        =   "Param19"
         UserName        =   "entregaexced"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P20 
         RealName        =   "Param20"
         UserName        =   "inicvigencia"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P21 
         RealName        =   "Param21"
         UserName        =   "usuario"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P22 
         RealName        =   "Param22"
         UserName        =   "datacad"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   16
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset169 
      CommandName     =   "Sel_TR02"
      CommDispId      =   2573
      RsDispId        =   2578
      CommandText     =   $"de_informaEM.dsx":31D9
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field8 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field10 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset170 
      CommandName     =   "Sel_TR02Codigo"
      CommDispId      =   2579
      RsDispId        =   2585
      CommandText     =   "select * from tb_tr02 where codigo = ? order by uf, cim, cidade, pesode"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   25
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "cim"
         Caption         =   "cim"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesode"
         Caption         =   "pesode"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesoate"
         Caption         =   "pesoate"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepeso"
         Caption         =   "fretepeso"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "porkilo"
         Caption         =   "porkilo"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalor"
         Caption         =   "fretevalor"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaate"
         Caption         =   "gen_txcoletaate"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletavalor"
         Caption         =   "gen_txcoletavalor"
      EndProperty
      BeginProperty Field18 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaexced"
         Caption         =   "gen_txcoletaexced"
      EndProperty
      BeginProperty Field19 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaate"
         Caption         =   "gen_txentregaate"
      EndProperty
      BeginProperty Field20 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregavalor"
         Caption         =   "gen_txentregavalor"
      EndProperty
      BeginProperty Field21 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaexced"
         Caption         =   "gen_txentregaexced"
      EndProperty
      BeginProperty Field22 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field23 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field25 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "codigo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset171 
      CommandName     =   "Sel_TR02UFCidadePeso"
      CommDispId      =   2586
      RsDispId        =   2599
      CommandText     =   "select * from tb_tr02 where status = '1' and codigo = ? and uf = ? and cidade = ? and pesode <= ? and pesoate >= ?"
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   25
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "cim"
         Caption         =   "cim"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesode"
         Caption         =   "pesode"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesoate"
         Caption         =   "pesoate"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepeso"
         Caption         =   "fretepeso"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "porkilo"
         Caption         =   "porkilo"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalor"
         Caption         =   "fretevalor"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaate"
         Caption         =   "gen_txcoletaate"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletavalor"
         Caption         =   "gen_txcoletavalor"
      EndProperty
      BeginProperty Field18 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaexced"
         Caption         =   "gen_txcoletaexced"
      EndProperty
      BeginProperty Field19 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaate"
         Caption         =   "gen_txentregaate"
      EndProperty
      BeginProperty Field20 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregavalor"
         Caption         =   "gen_txentregavalor"
      EndProperty
      BeginProperty Field21 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaexced"
         Caption         =   "gen_txentregaexced"
      EndProperty
      BeginProperty Field22 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field23 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field25 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   5
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "cod"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "cidade"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   35
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "pesode"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "pesoate"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset172 
      CommandName     =   "Sel_TR02UFCimPeso"
      CommDispId      =   2593
      RsDispId        =   2606
      CommandText     =   "select * from tb_tr02 where status = '1' and codigo = ? and uf = ? and cim = ? and pesode <= ? and pesoate >= ? "
      ActiveConnectionName=   "cn_informa"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   25
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "status"
         Caption         =   "status"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   200
         Name            =   "statusdescr"
         Caption         =   "statusdescr"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "codigo"
         Caption         =   "codigo"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "origem"
         Caption         =   "origem"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "origemuf"
         Caption         =   "origemuf"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "descricao"
         Caption         =   "descricao"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "uf"
         Caption         =   "uf"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "cim"
         Caption         =   "cim"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   35
         Scale           =   0
         Type            =   200
         Name            =   "cidade"
         Caption         =   "cidade"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesode"
         Caption         =   "pesode"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "pesoate"
         Caption         =   "pesoate"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretepeso"
         Caption         =   "fretepeso"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "porkilo"
         Caption         =   "porkilo"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "complemento"
         Caption         =   "complemento"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "fretevalor"
         Caption         =   "fretevalor"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaate"
         Caption         =   "gen_txcoletaate"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletavalor"
         Caption         =   "gen_txcoletavalor"
      EndProperty
      BeginProperty Field18 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txcoletaexced"
         Caption         =   "gen_txcoletaexced"
      EndProperty
      BeginProperty Field19 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaate"
         Caption         =   "gen_txentregaate"
      EndProperty
      BeginProperty Field20 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregavalor"
         Caption         =   "gen_txentregavalor"
      EndProperty
      BeginProperty Field21 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "gen_txentregaexced"
         Caption         =   "gen_txentregaexced"
      EndProperty
      BeginProperty Field22 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "inicvigencia"
         Caption         =   "inicvigencia"
      EndProperty
      BeginProperty Field23 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "fimvigencia"
         Caption         =   "fimvigencia"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "usuariocad"
         Caption         =   "usuariocad"
      EndProperty
      BeginProperty Field25 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "datacad"
         Caption         =   "datacad"
      EndProperty
      NumGroups       =   0
      ParamCount      =   5
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "cod"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   7
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Param2"
         UserName        =   "uf"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Param3"
         UserName        =   "cim"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Param4"
         UserName        =   "pesode"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Param5"
         UserName        =   "pesoate"
         Direction       =   1
         Precision       =   19
         Scale           =   4
         Size            =   8
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "de_informaEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub rsSel_CadCidadePorCidade_WillChangeField(ByVal cFields As Long, ByVal Fields As Variant, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

