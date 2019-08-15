unit MSJet.API;

{$WARN UNSAFE_TYPE OFF}
{$WARN UNSAFE_CODE OFF}

interface

uses Winapi.Windows, System.SysUtils;

type
  EMSJetError = class(Exception);

  JET_API_PTR  = Cardinal;
  JET_INSTANCE = JET_API_PTR;
  JET_ERR      = Integer;
  JET_SESID    = JET_API_PTR;
  JET_PCSTR    = PAnsiChar;
  JET_PCWSTR   = PWideChar;
  JET_DBID     = ULONG;
  JET_GRBIT    = ULONG;
  JET_TABLEID  = JET_API_PTR;
  JET_COLTYP   = ULONG;
  JET_COLUMNID = ULONG;
  JET_OBJTYP   = ULONG;

type
  tag_JET_COLUMNCREATE = record
      cbStruct: ULONG;
      szColumnName: JET_PCSTR;
      coltyp: JET_COLTYP;
      cbMax: ULONG;
      grbit: JET_GRBIT;
      pvDefault: Pointer;
      cbDefault: ULONG;
      cp: ULONG;
      columnid: JET_COLUMNID;
      err: JET_ERR;
  end;
  PJET_COLUMNCREATE = ^JET_COLUMNCREATE;
  JET_COLUMNCREATE = tag_JET_COLUMNCREATE;

  tagJET_UNICODEINDEX = record
    lcid: ULONG;
    dwMapFlags: ULONG;
  end;
  PJET_UNICODEINDEX = ^JET_UNICODEINDEX;
  JET_UNICODEINDEX = tagJET_UNICODEINDEX;

  tagJET_TUPLELIMITS = record
    chLengthMin: ULONG;
    chLengthMax: ULONG;
    chToIndexMax: ULONG;
    cchIncrement: ULONG;
    ichStart: ULONG;
  end;
  PJET_TUPLELIMITS = ^JET_TUPLELIMITS;
  JET_TUPLELIMITS  = tagJET_TUPLELIMITS;

  tagJET_CONDITIONALCOLUMN = record
    cbStruct: ULONG;
    szColumnName: JET_PCSTR;
    grbit: JET_GRBIT;
  end;
  PJET_CONDITIONALCOLUMN = ^JET_CONDITIONALCOLUMN;
  JET_CONDITIONALCOLUMN = tagJET_CONDITIONALCOLUMN;

  tagJET_INDEXCREATE = record
    cbStruct: ULONG;
    szIndexName: JET_PCSTR;
    szKey: JET_PCSTR;
    cbKey: ULONG;
    grbit: JET_GRBIT;
    ulDensity: ULONG;
    lcid: ULONG;
    UNICODE: ULONG;
    rgconditionalcolumn: PJET_CONDITIONALCOLUMN;
    cConditionalColumn: ULONG;
    err: JET_ERR;
    cbKeyMost: ULONG;
  end;

  PJET_INDEXCREATE = ^JET_INDEXCREATE;
  JET_INDEXCREATE = tagJET_INDEXCREATE;

  tagJET_TABLECREATE = record
    cbStruct: ULONG;
    szTableName: JET_PCSTR;
    szTemplateTableName: JET_PCSTR;
    ulPages: ULONG;
    ulDensity: ULONG;
    rgcolumncreate: PJET_COLUMNCREATE;
    cColumns: ULONG;
    rgindexcreate: PJET_INDEXCREATE;
    cIndexes: ULONG;
    grbit: JET_GRBIT;
    tableid: JET_TABLEID;
    cCreated: ULONG;
  end;
  PJET_TABLECREATE = ^JET_TABLECREATE;
  JET_TABLECREATE = tagJET_TABLECREATE;

  tagJET_TABLECREATE2 = record
    cbStruct: ULONG;
    szTableName: JET_PCSTR;
    szTemplateTableName: JET_PCSTR;
    ulPages: ULONG;
    ulDensity: ULONG;
    rgcolumncreate: PJET_COLUMNCREATE;
    cColumns: ULONG;
    rgindexcreate: PJET_INDEXCREATE;
    cIndexes: ULONG;
    szCallback: JET_PCSTR;
    cbtyp: ULONG;
    grbit: JET_GRBIT;
    tableid: JET_TABLEID;
    cCreated: ULONG;
  end;
  PJET_TABLECREATE2 = ^JET_TABLECREATE2;
  JET_TABLECREATE2 = tagJET_TABLECREATE2;

  JET_COLUMNDEF = record
    cbStruct: ULONG;
    columnid: JET_COLUMNID;
    coltyp: JET_COLTYP;
    wCountry: Word;
    langid: Word;
    cp: Word;
    wCollate: Word;
    cbMax: ULONG;
    grbit: JET_GRBIT;
  end;

  PJET_COLUMNDEF = ^JET_COLUMNDEF;

  JET_COLUMNBASE = record
    cbStruct: ULONG;
    columnid: JET_COLUMNID;
    coltyp: JET_COLTYP;
    wCountry: Word;
    langid: Word;
    cp: Word;
    wFiller: Word;
    cbMax: ULONG;
    grbit: JET_GRBIT;
    szBaseTableName: array[0..255] of AnsiChar;
    szBaseColumnName: array[0..255] of AnsiChar;
  end;

  JET_COLUMNLIST = record
    cbStruct: ULONG;
    tableid: JET_TABLEID;
    cRecord: ULONG;
    PresentationOrder: JET_COLUMNID;
    columnname: JET_COLUMNID;
    columnid: JET_COLUMNID;
    coltyp: JET_COLUMNID;
    Country: JET_COLUMNID;
    Langid: JET_COLUMNID;
    Cp: JET_COLUMNID;
    Collate: JET_COLUMNID;
    cbMax: JET_COLUMNID;
    grbit: JET_COLUMNID;
    Default: JET_COLUMNID;
    BaseTableName: JET_COLUMNID;
    BaseColumnName: JET_COLUMNID;
    DefinitionName: JET_COLUMNID;
  end;

  JET_INDEXLIST = record
    cbStruct: ULONG;
    tableid: JET_TABLEID;
    cRecord: ULONG;
    columnidindexname: JET_COLUMNID;
    columnidgrbitIndex: JET_COLUMNID;
    columnidcKey: JET_COLUMNID;
    columnidcEntry: JET_COLUMNID;
    columnidcPage: JET_COLUMNID;
    columnidcColumn: JET_COLUMNID;
    columnidiColumn: JET_COLUMNID;
    columnidcolumnid: JET_COLUMNID;
    columnidcoltyp: JET_COLUMNID;
    columnidCountry: JET_COLUMNID;
    columnidLangid: JET_COLUMNID;
    columnidCp: JET_COLUMNID;
    columnidCollate: JET_COLUMNID;
    columnidgrbitColumn: JET_COLUMNID;
    columnidcolumnname: JET_COLUMNID;
    columnidLCMapFlags: JET_COLUMNID;
  end;
  PJET_INDEXLIST = ^JET_INDEXLIST;

  JET_RETINFO = record
    cbStruct: ULONG;
    ibLongValue: ULONG;
    itagSequence: ULONG;
    columnidNextTagged: JET_COLUMNID;
  end;
  PJET_RETINFO = ^JET_RETINFO;

  JET_SETINFO = record
    cbStruct: ULONG;
    ibLongValue: ULONG;
    itagSequence: ULONG;
  end;
  PJET_SETINFO = ^JET_SETINFO;

  JET_OBJECTLIST = record
    cbStruct: ULONG;
    tableid: JET_TABLEID;
    cRecord: ULONG;
    columnidcontainername: JET_COLUMNID;
    columnidobjectname: JET_COLUMNID;
    columnidobjtyp: JET_COLUMNID;
    columniddtCreate: JET_COLUMNID;
    columniddtUpdate: JET_COLUMNID;
    columnidgrbit: JET_COLUMNID;
    columnidflags: JET_COLUMNID;
    columnidcRecord: JET_COLUMNID;
    columnidcPage: JET_COLUMNID;
  end;
  PJET_OBJECTLIST = ^JET_OBJECTLIST;

  JET_RECPOS = record
    cbStruct: ULONG;
    centriesLT: ULONG;
    centriesInRange: ULONG;
    centriesTotal: ULONG;
  end;
  PJET_RECPOS = ^JET_RECPOS;

function JetError(Err: JET_ERR; instance: JET_INSTANCE = 0; sesid: JET_SESID = 0): JET_ERR; inline;

function JetAddColumn(sesid: JET_SESID; tableid: JET_TABLEID; szColumnName: JET_PCSTR; pcolumndef: PJET_COLUMNDEF; pvDefault: Pointer; cbDefault: ULONG; var columnid: JET_COLUMNID): JET_ERR; stdcall;
function JetAttachDatabase(sesid: JET_SESID; szFilename: JET_PCSTR; grbit: JET_GRBIT): JET_ERR; stdcall;
function JetBeginSession(pinstance: JET_INSTANCE; var sesid: JET_SESID; szUserName, szPassword: JET_PCSTR): JET_ERR; stdcall;
function JetBeginTransaction(sesid: JET_SESID): JET_ERR; stdcall;
function JetCloseDatabase(sesid: JET_SESID; pdbid: JET_DBID; grbit: JET_GRBIT): JET_ERR; stdcall;
function JetCloseTable(sesid: JET_SESID; ptableid: JET_TABLEID): JET_ERR; stdcall;
function JetCommitTransaction(sesid: JET_SESID; grbit: JET_GRBIT): JET_ERR; stdcall;
function JetCreateDatabase(sesid: JET_SESID; szFilename, szConnect: JET_PCSTR; var pdbid: JET_DBID; grbit: JET_GRBIT): JET_ERR; stdcall;
function JetCreateIndex(sesid: JET_SESID; tableid: JET_TABLEID; szIndexName: JET_PCSTR; grbit: JET_GRBIT; szKey: JET_PCSTR; cbKey: ULONG; lDensity: ULONG): JET_ERR; stdcall;
function JetCreateInstance(var pinstance: JET_INSTANCE; const InstanceName: PAnsiChar): JET_ERR; stdcall;
function JetCreateTable(sesid: JET_SESID; dbid: JET_DBID; szTableName: JET_PCSTR; lPages, lDensity: ULONG; var ptableid: JET_TABLEID): JET_ERR; stdcall;
function JetCreateTableColumnIndex(sesid: JET_SESID; dbid: JET_DBID; var ptablecreate: JET_TABLECREATE): JET_ERR; stdcall;
function JetDelete(sesid: JET_SESID; tableid: JET_TABLEID): JET_ERR; stdcall;
function JetDeleteColumn(sesid: JET_SESID; tableid: JET_TABLEID; szColumnName: PAnsiChar): JET_ERR; stdcall;
function JetDeleteIndex(sesid: JET_SESID; tableid: JET_TABLEID; szIndexName: PAnsiChar): JET_ERR; stdcall;
function JetDeleteTable(sesid: JET_SESID; dbid: JET_DBID; szTableName: PAnsiChar): JET_ERR; stdcall;
function JetDetachDatabase(sesid: JET_SESID; szFilename: JET_PCSTR): JET_ERR; stdcall;
function JetEndSession(sesid: JET_SESID; grbit: JET_GRBIT): JET_ERR; stdcall;
function JetGetObjectInfo(sesid: JET_SESID; dbid: JET_DBID; objtyp: JET_OBJTYP; szContainerName: JET_PCSTR; szObjectName: JET_PCSTR; pvResult: Pointer; cbMax: ULONG; InfoLevel: ULONG): JET_ERR; stdcall;
function JetGetSessionParameter(sesid: JET_SESID; sesparamid: ULONG; pvParam: Pointer; cbParamMax: ULONG; var pcbParamActual): JET_ERR; stdcall;
function JetGetSystemParameter(instance: JET_INSTANCE; sesid: JET_SESID; paramid: ULONG; var plParam: JET_API_PTR; szParam: JET_PCSTR; cbMax: ULONG): JET_ERR; stdcall;
function JetGetTableColumnInfo(sesid: JET_SESID; tableid: JET_TABLEID; szColumnName: JET_PCSTR; pvResult: Pointer; cbMax, InfoLevel: ULONG): JET_ERR; stdcall;
function JetGetTableIndexInfo(sesid: JET_SESID; tableid: JET_TABLEID; szIndexName: JET_PCSTR; pvResult: Pointer; cbResult, InfoLevel: ULONG): JET_ERR; stdcall;
function JetInit(var pinstance: JET_INSTANCE): JET_ERR; stdcall;
function JetMove(sesid: JET_SESID; tableid: JET_TABLEID; cRow: Integer; grbit: JET_GRBIT): JET_ERR; stdcall;
function JetOpenDatabase(sesid: JET_SESID; szFilename, szConnect: JET_PCSTR; var pdbid: JET_DBID; grbit: JET_GRBIT): JET_ERR; stdcall;
function JetOpenTable(sesid: JET_SESID; dbid: JET_DBID; szTableName: PAnsiChar; pvParameters: Pointer; cbParameters: ULONG; grbit: JET_GRBIT; var ptableid: JET_TABLEID): JET_ERR; stdcall;
function JetPrepareUpdate(sesid: JET_SESID; tableid: JET_TABLEID; prep: ULONG): JET_ERR; stdcall;
function JetRenameColumn(sesid: JET_SESID; tableid: JET_TABLEID; szName, szNameNew: JET_PCSTR; grbit: JET_GRBIT): JET_ERR; stdcall;
function JetRenameTable(sesid: JET_SESID; dbid: JET_DBID; szName, szNameNew: JET_PCSTR): JET_ERR; stdcall;
function JetRetrieveColumn(sesid: JET_SESID; tableid: JET_TABLEID; columnid: JET_COLUMNID; pvData: Pointer; cbData: ULONG; var pcbActual: ULONG; grbit: JET_GRBIT; pretinfo: PJET_RETINFO): JET_ERR; stdcall;
function JetRollback(sesid: JET_SESID; grbit: JET_GRBIT): JET_ERR; stdcall;
function JetSetColumn(sesid: JET_SESID; tableid: JET_TABLEID; columnid: JET_COLUMNID; pvData: Pointer; cbData: ULONG; grbit: JET_GRBIT; psetinfo: PJET_SETINFO): JET_ERR; stdcall;
function JetSetCurrentIndex(sesid: JET_SESID; tableid: JET_TABLEID; szIndexName: PAnsiChar): JET_ERR; stdcall;
function JetSetSystemParameter(var pinstance: JET_INSTANCE; sesid: JET_SESID; paramid: ULONG; lParam: JET_API_PTR; szParam: JET_PCSTR): JET_ERR; stdcall;
function JetTerm(pinstance: JET_INSTANCE): JET_ERR; stdcall;
function JetUpdate(sesid: JET_SESID; tableid: JET_TABLEID; pvBookmark: Pointer; cbBookmark: ULONG; var pcbActual: ULONG): JET_ERR; stdcall;
function JetMakeKey(sesid: JET_SESID; tableid: JET_TABLEID; pvData: Pointer; cbData: ULONG; grdibt: JET_GRBIT): JET_ERR; stdcall;
function JetSeek(sesid: JET_SESID; tableid: JET_TABLEID; grdibt: JET_GRBIT): JET_ERR; stdcall;
function JetRetrieveKey(sesid: JET_SESID; tableid: JET_TABLEID; pvData: Pointer; cbMax: ULONG; var pcbActual: ULONG; grdibt: JET_GRBIT): JET_ERR; stdcall;
function JetSetIndexRange(sesid: JET_SESID; tableid: JET_TABLEID; grdibt: JET_GRBIT): JET_ERR; stdcall;
function JetGetRecordPosition(sesid: JET_SESID; tableid: JET_TABLEID; precpos: PJET_RECPOS; cbRecpos: ULONG): JET_ERR; stdcall;
function JetGotoPosition(sesid: JET_SESID; tableid: JET_TABLEID; precpos: PJET_RECPOS): JET_ERR; stdcall;

function JetGetBookmark(sesid: JET_SESID; tableid: JET_TABLEID; pvBookmark: Pointer; cbMax: ULONG; var pcbActual: ULONG): JET_ERR; stdcall;
function JetGotoBookmark(sesid: JET_SESID; tableid: JET_TABLEID; pvBookmark: Pointer; cbBookmark: ULONG): JET_ERR; stdcall;

function JetGetSecondaryIndexBookmark(sesid: JET_SESID; tableid: JET_TABLEID;
                                      pvSecondaryKey: Pointer; cbSecondaryKeyMax: ULONG; var pcbSecondaryKeyActual: ULONG;
                                      pvPrimaryKey: Pointer; cbPrimaryKeyMax: ULONG; var pcbPrimaryKeyActual: ULONG;
                                      grdibt: JET_GRBIT): JET_ERR; stdcall;

function JetGotoSecondaryIndexBookmark(sesid: JET_SESID; tableid: JET_TABLEID;
                                      pvSecondaryKey: Pointer; cbSecondaryKey: ULONG;
                                      pvPrimaryKey: Pointer; cbPrimaryKey: ULONG;
                                      grdibt: JET_GRBIT): JET_ERR; stdcall;

implementation

const
  esent = 'esent.dll';

function JetAddColumn; external esent name 'JetAddColumn';
function JetAttachDatabase; external esent name 'JetAttachDatabase';
function JetBeginSession; external esent name 'JetBeginSession';
function JetBeginTransaction; external esent name 'JetBeginTransaction';
function JetCloseDatabase; external esent name 'JetCloseDatabase';
function JetCloseTable; external esent name 'JetCloseTable';
function JetCommitTransaction; external esent name 'JetCommitTransaction';
function JetCreateDatabase; external esent name 'JetCreateDatabase';
function JetCreateIndex; external esent name 'JetCreateIndex';
function JetCreateInstance; external esent name 'JetCreateInstance';
function JetCreateTable; external esent name 'JetCreateTable';
function JetCreateTableColumnIndex; external esent name 'JetCreateTableColumnIndex';
function JetDelete; external esent name 'JetDelete';
function JetDeleteColumn; external esent name 'JetDeleteColumn';
function JetDeleteIndex; external esent name 'JetDeleteIndex';
function JetDeleteTable; external esent name 'JetDeleteTable';
function JetDetachDatabase; external esent name 'JetDetachDatabase';
function JetEndSession; external esent name 'JetEndSession';
function JetGetObjectInfo; external esent name 'JetGetObjectInfo';
function JetGetSessionParameter; external esent name 'JetGetSessionParameter';
function JetGetSystemParameter; external esent name 'JetGetSystemParameter';
function JetGetTableColumnInfo; external esent name 'JetGetTableColumnInfo';
function JetGetTableIndexInfo; external esent name 'JetGetTableIndexInfo';
function JetInit; external esent name 'JetInit';
function JetMove; external esent name 'JetMove';
function JetOpenDatabase; external esent name 'JetOpenDatabase';
function JetOpenTable; external esent name 'JetOpenTable';
function JetPrepareUpdate; external esent name 'JetPrepareUpdate';
function JetRenameColumn; external esent name 'JetRenameColumn';
function JetRenameTable; external esent name 'JetRenameTable';
function JetRetrieveColumn; external esent name 'JetRetrieveColumn';
function JetRollback; external esent name 'JetRollback';
function JetSetColumn; external esent name 'JetSetColumn';
function JetSetCurrentIndex; external esent name 'JetSetCurrentIndex';
function JetSetSystemParameter; external esent name 'JetSetSystemParameter';
function JetTerm; external esent name 'JetTerm';
function JetUpdate; external esent name 'JetUpdate';
function JetMakeKey; external esent name 'JetMakeKey';
function JetSeek; external esent name 'JetSeek';
function JetRetrieveKey; external esent name 'JetRetrieveKey';
function JetSetIndexRange; external esent name 'JetSetIndexRange';
function JetGetRecordPosition; external esent name 'JetGetRecordPosition';
function JetGotoPosition; external esent name 'JetGotoPosition';
function JetGetBookmark; external esent name 'JetGetBookmark';
function JetGotoBookmark; external esent name 'JetGotoBookmark';
function JetGetSecondaryIndexBookmark; external esent name 'JetGetSecondaryIndexBookmark';
function JetGotoSecondaryIndexBookmark; external esent name 'JetGotoSecondaryIndexBookmark';

function  JetError(Err: JET_ERR; instance: JET_INSTANCE = 0; sesid: JET_SESID = 0): JET_ERR;
var
  Code: ULONG;
  Buf: array[0..1024] of AnsiChar;
begin
  Result := Err;
  if Err < 0 then
  begin
    Code := Err;
    JetGetSystemParameter(instance, sesid, 70, Code, PAnsiChar(@Buf[0]), 1024);
    raise EMSJetError.Create(string(AnsiString(Buf)));
  end;
end;

end.
