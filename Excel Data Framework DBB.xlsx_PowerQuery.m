// Power Query from: Excel Data Framework DBB.xlsx
// Pathname: c:\Users\daniel-bighelini\OneDrive\Documentos\Planilhas\Excel Data Framework DBB\Excel Data Framework DBB.xlsx
// Extracted: 2026-07-26T16:56:15.506Z

section Section1;

shared srcObjetosPowerQuery = let

    Resultado =

        Record.FieldNames(

            #sections[Section1]

        )

in

    Resultado
;

shared srcCategoriasPowerQuery = let
    Fonte = srcWorkbook{[Name=parTabelaCategoriasConsultasPQ]}[Content]
in
    Fonte;

shared srcWorkbook = let
    srcWorkbook = Table.Buffer(Excel.CurrentWorkbook())
in
    srcWorkbook;

shared srcParametrosExcel = let
    // Procura a tabela definida no parâmetro 'parTabelaParametros' e, caso não exista,
    // retorna uma tabela vazia para evitar erros nas consultas dependentes.

    Fonte =
        try
            srcWorkbook{[Name = parTabelaParametros]}[Content]
            otherwise
                #table(
                    type table
                    [
                        Parâmetro = text,
                        Valor = any
                    ],
                    {}
                )
in
    Fonte;

shared srcParametrosFormatosArquivos = let
    Fonte = srcWorkbook{[Name=parTabelaParametrosFormatosArquivos]}[Content]
in
    Fonte;

shared srcParametrosTratamentos = let
    Fonte = srcWorkbook{[Name=parTabelaParametrosTratamentos]}[Content]
in
    Fonte;

shared srcParametrosValidacoes = let
    Fonte = srcWorkbook{[Name=parTabelaParametrosValidacoes]}[Content]
in
    Fonte;

shared srcParametrosSeveridades = let
    Fonte = srcWorkbook{[Name=parTabelaParametrosSeveridades]}[Content]
in
    Fonte;

shared srcSchema = let
    Fonte = srcWorkbook{[Name=parTabelaSchema]}[Content]
in
    Fonte;

shared srcClientes = let
    Fonte = srcWorkbook{[Name=parTabelaClientes]}[Content]
in
    Fonte;

shared srcProdutos = let
    Fonte = srcWorkbook{[Name=parTabelaProdutos]}[Content]
in
    Fonte;

shared srcVendas = let
    Fonte = srcWorkbook{[Name=parTabelaVendas]}[Content]
in
    Fonte;

shared srcDadosGenericos = let
    Fonte = srcWorkbook{[Name=parTabelaDadosGenericos]}[Content]
in
    Fonte;

shared parTabelaParametros = "tbParametros" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared cfgParametros = // Materializa a tabela consolidada em um Record para permitir
// acesso O(1) aos parâmetros através de Record.Field.
Record.FromTable(
    Table.Buffer(stgParametros)
);

shared cfgCategoriasPowerQuery = let
    Fonte =
        Table.Buffer(
            stgCategoriasPowerQuery
        ),

    Registros =
        Table.AddColumn(
            Fonte,
            "Value",
            each
                [
                    Ordem = [Ordem],
                    Categoria = [Categoria],
                    Objetivo = [Objetivo],
                    Saída = [Saída]
                ],
            type record
        ),

    Resultado =
        Record.FromTable(
            Table.RenameColumns(
                Table.SelectColumns(
                    Registros,
                    {
                        "Prefixo",
                        "Value"
                    }
                ),
                {
                    {"Prefixo", "Name"}
                }
            )
        )
in
    Resultado;

shared fxParametro = (
    parametro as text,
    optional valorPadrao as nullable any
) as any =>

let
    Parametros =
        cfgParametros,

    Parametro =
        Record.FieldOrDefault(
            Parametros,
            parametro
        ),

    Tipo =
        if Parametro = null then
            type any
        else
            Record.FieldOrDefault(
                Parametro,
                "Tipo",
                type any
            ),

    Obrigatorio =
        if Parametro = null then
            true
        else
            Record.FieldOrDefault(
                Parametro,
                "Obrigatório",
                true
            ),

    Permitidos =
        if Parametro = null then
            null
        else
            Record.FieldOrDefault(
                Parametro,
                "Permitidos"
            ),

    Local =
        "pt-BR",

    Valor =
        if Parametro = null then
            valorPadrao
        else
            Parametro[Valor],

    Resultado =
        if Parametro = null and Obrigatorio and valorPadrao = null then

            error Error.Record(
                "Parâmetro inexistente",
                Text.Format(
                    "O parâmetro '#{0}' não foi encontrado.",
                    {parametro}
                ),
                [
                    Parametro = parametro,
                    Disponiveis =
                        Text.Combine(
                            List.Sort(
                                Record.FieldNames(Parametros)
                            ),
                            ", "
                        )
                ]
            )

        else if Valor = null then

            null

        else

            let
                Conversao =
                    try
                        fxConversor(
                            Valor,
                            Tipo,
                            Local
                        )
            in
                if Conversao[HasError] then

                    error Error.Record(
                        "Erro de conversão",
                        Text.Format(
                            "Não foi possível converter o valor '#{0}' do parâmetro '#{1}'.",
                            {
                                Valor,
                                parametro
                            }
                        ),
                        [
                            Parametro = parametro,
                            Valor = Valor,
                            TipoDestino = Tipo
                        ]
                    )

                else

                    let
                        ValorConvertido =
                            Conversao[Value],

                        ValorPermitido =
                            if Permitidos = null then
                                true
                            else
                                List.Contains(
                                    Permitidos,
                                    ValorConvertido
                                )

                    in
                        if not ValorPermitido then

                            error Error.Record(
                                "Valor inválido",
                                Text.Format(
                                    "O valor '#{0}' não é permitido para o parâmetro '#{1}'.",
                                    {
                                        ValorConvertido,
                                        parametro
                                    }
                                ),
                                [
                                    Parametro = parametro,
                                    Valor = ValorConvertido,
                                    Permitidos = Permitidos
                                ]
                            )

                        else

                            ValorConvertido

in
    Resultado;

shared stgCategoriasPowerQuery = let

    Fonte =
        srcCategoriasPowerQuery,

    LinhasValidas = Table.SelectRows(
        Fonte,
        each [Prefixo] <> null and Text.Trim(Text.From([Prefixo])) <> "" and
             [Categoria] <> null and Text.Trim(Text.From([Categoria])) <> ""
    ),

    Distintos =
        Table.Distinct(
            LinhasValidas,
            {"Prefixo"}
        ),
    TipoAlterado = Table.TransformColumnTypes(Distintos,{{"Ordem", Int64.Type}, {"Prefixo", type text}, {"Categoria", type text}, {"Objetivo", type text}, {"Saída", type text}})

in

    TipoAlterado;

shared stgObjetosExcel = /*
------------------------------------------------------------------------------
Consulta.....: stgObjetosExcel

Descrição....:
Catálogo normalizado dos objetos do Excel.

Cada registro representa um objeto do Workbook juntamente com seus
metadados identificados pelo framework.

Esta consulta não cria índices nem estruturas hierárquicas.
------------------------------------------------------------------------------
*/

let

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

    Fonte =

        Table.Column(
            srcWorkbook,
            "Name"
        ),

//--------------------------------------------------------------------------
// Objetos
//--------------------------------------------------------------------------

    Objetos =

        List.Transform(

            Fonte,

            (Nome) =>

                Record.Combine(

                    {

                        [

                            Origem = "Excel",

                            Nome = Nome

                        ],

                        fxIdentificarObjetoPeloNome(

                            Nome,

                            "Excel"

                        )

                    }

                )

        ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        Table.FromRecords(

            Objetos

        )

in

    Resultado;

shared stgObjetosPowerQuery = /*
------------------------------------------------------------------------------
Consulta.....: stgObjetosPowerQuery

Descrição....:
Catálogo normalizado dos objetos do Power Query.

Cada registro representa um objeto existente no ambiente Power Query
juntamente com seus metadados identificados pelo framework.

Esta consulta não cria índices nem estruturas hierárquicas.
------------------------------------------------------------------------------
*/

let

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

    Fonte =

        srcObjetosPowerQuery,

//--------------------------------------------------------------------------
// Objetos
//--------------------------------------------------------------------------

    Objetos =

            List.Transform(

                Fonte,

                (Nome) =>

                    let

                        Metadados =

                                fxIdentificarObjetoPeloNome(

                                    Nome,

                                    "PowerQuery"

                                )

                    in

                        if Metadados = null then
                            null
                        else

                        Record.Combine(

                            {

                                [

                                    Origem = "PowerQuery",

                                    Nome = Nome

                                ],

                                Metadados

                            }

                        )

        ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        Table.FromRecords(

            Objetos

        )

in

    Resultado;

shared stgObjetos = /*
------------------------------------------------------------------------------
Consulta.....: stgObjetos

Descrição....:
Consolida o catálogo de objetos provenientes de todas as origens do framework.

Responsabilidades:

    • Unificar os objetos do Power Query.
    • Unificar os objetos do Excel.
    • Não criar índices.
    • Não alterar metadados.

A indexação é realizada posteriormente por cfgObjetos.

------------------------------------------------------------------------------
*/

let

//--------------------------------------------------------------------------
// Fontes
//--------------------------------------------------------------------------

    Fonte =

        Table.Combine(

            {

                stgObjetosPowerQuery,

                stgObjetosExcel

            }

        ),

//--------------------------------------------------------------------------
// Ordenação
//--------------------------------------------------------------------------

    Resultado =

        Table.Sort(

            Fonte,

            {

                {"Origem", Order.Ascending},
                {"Nome", Order.Ascending}

            }

        )

in

    Resultado;

shared stgTabelasExcel = /*
------------------------------------------------------------------------------
Consulta.....: stgTabelasExcel

Descrição....:
Catálogo normalizado das tabelas do Excel.

Cada registro representa uma tabela do Workbook juntamente com seus
metadados estruturais.

Não materializa nem armazena o conteúdo das tabelas.

------------------------------------------------------------------------------
*/

let

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

    Fonte =

        stgObjetosExcel,

//--------------------------------------------------------------------------
// Metadados
//--------------------------------------------------------------------------

    Registros =

        List.Transform(

            Table.ToRecords(
                Fonte
            ),

            (Objeto) =>

                let

                    Tabela =

                        srcWorkbook{[Name = Objeto[Nome]]}[Content]

                in

                    Record.Combine(

                        {

                            Objeto,

                            [

                                Columns = Table.ColumnNames(Tabela),

                                ColumnCount = Table.ColumnCount(Tabela)

                            ]

                        }

                    )

        ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        Table.FromRecords(
            Registros
        )

in

    Resultado;

shared stgTabelasPowerQuery = let

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

Fonte =

    Table.SelectRows(

        stgObjetosPowerQuery,

        each
            [Kind] = "Table"
            and
            [Nome] <> "srcSections" and
            [Nome] <> "stgSections" and
            [Nome] <> "stgObjetos" and
            [Nome] <> "stgObjetosPowerQuery" and
            [Nome] <> "cfgObjetos" and
            [Nome] <> "cfgObjetosPowerQuery" and
            [Nome] <> "stgTabelas" and
            [Nome] <> "cfgTabelas" and
            [Nome] <> "srcTabelasPowerQuery" and
            [Nome] <> "stgTabelasPowerQuery" and
            [Nome] <> "cfgTabelasPowerQuery"

    ),

//--------------------------------------------------------------------------
// Metadados
//--------------------------------------------------------------------------

    Registros =

        List.Transform(

            Table.ToRecords(
                Fonte
            ),

            (Objeto) =>

                let

                    Tabela =

                        Record.Field(

                            #sections[Section1],

                            Objeto[Nome]

                        ),

                    Metadados =

                        /*
                        if Value.Is(Tabela, type table) then

                            [

                                Columns = Table.ColumnNames(Tabela),
                                ColumnCount = Table.ColumnCount(Tabela)

                            ]

                        else if Value.Is(Tabela, type list) then

                            let

                                Primeiro = List.First(Tabela, null)

                            in

                                if Value.Is(Primeiro, type record) then

                                    [

                                        Columns = Record.FieldNames(Primeiro),
                                        ColumnCount = List.Count(Record.FieldNames(Primeiro))

                                    ]

                                else

                                    [

                                        Columns = {},
                                        ColumnCount = 0

                                    ]

                        else
*/
                            [

                                Columns = {},
                                ColumnCount = 0

                            ]


                in

                    Record.Combine({
                        Objeto,
                        Metadados
                    })

        ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        Table.FromRecords(

            Registros

        )

in

    Resultado;

shared stgTabelas = let

    Resultado =

        Table.Combine({

            stgTabelasPowerQuery,

            stgTabelasExcel

        })

in

    Resultado;

shared stgParametrosExcel = let

    Fonte =
        Table.SelectColumns(
            srcParametrosExcel,
            {
                "Parâmetro",
                "Valor",
                "Tipo",
                "Obrigatório",
                "Permitidos"
            }
        ),

    LinhasValidas =
        Table.SelectRows(
            Fonte,
            each
                [Parâmetro] <> null
                    and Text.Trim(Text.From([Parâmetro])) <> ""
        ),

    Padronizar =
        Table.TransformColumns(
            LinhasValidas,
            {
                {
                    "Parâmetro",
                    each Text.Trim(Text.From(_)),
                    type text
                },
                {
                    "Tipo",
                    each fxParametroTipo(_),
                    type nullable type
                },
                {
                    "Obrigatório",
                    each
                        fxConversor(
                            _,
                            type logical
                        ),
                    type nullable logical
                }
            }
        ),

    ConverterPermitidos =
        Table.AddColumn(
            Padronizar,
            "Permitidos2",
            each
                if [Permitidos] = null then
                    null
                else
                    List.Transform(
                        fxLista([Permitidos]),
                        (v) =>
                            fxConversor(
                                v,
                                [Tipo]
                            )
                    ),
            type nullable list
        ),

    RemoverPermitidosOriginal =
        Table.RemoveColumns(
            ConverterPermitidos,
            {"Permitidos"}
        ),

    RenomearPermitidos =
        Table.RenameColumns(
            RemoverPermitidosOriginal,
            {
                {
                    "Permitidos2",
                    "Permitidos"
                }
            }
        ),

    AdicionarOrigem =
        Table.AddColumn(
            RenomearPermitidos,
            "Origem",
            each "Excel",
            type text
        )

in

    AdicionarOrigem;

shared stgParametrosPowerQuery = let

    Fonte =
        #sections[Section1],

    PrefixoParametro =
        List.First(
            List.Select(
                Record.FieldNames(cfgCategoriasPowerQuery),
                each
                    Record.Field(
                        cfgCategoriasPowerQuery,
                        _
                    )[Categoria] = "Parâmetros"
            )
        ),

    NomesCandidatos =
        List.Select(
            Record.FieldNames(Fonte),
            each
                Text.StartsWith(
                    _,
                    PrefixoParametro
                )
        ),

    Parametros =
        Table.FromRecords(
            List.RemoveNulls(
                List.Transform(
                    NomesCandidatos,
                    (Nome) =>

                        let

                            Definicao =
                                fxParametroLerParametroPQ(
                                    Record.Field(
                                        Fonte,
                                        Nome
                                    )
                                )

                        in

                            if Definicao = null then
                                null
                            else
                                Record.Combine(
                                    {
                                        [
                                            Parâmetro = Nome,
                                            Origem = "PowerQuery"
                                        ],
                                        Definicao
                                    }
                                )
                )
            ),
            type table
            [
                Parâmetro = text,
                Valor = any,
                Tipo = nullable type,
                Obrigatório = nullable logical,
                Permitidos = nullable list,
                Origem = text
            ]
        )

in

    Parametros;

shared stgParametros = let

    Fonte =
        Table.Combine(
            {
                stgParametrosExcel,
                stgParametrosPowerQuery
            }
        ),

    Distintos =
        Table.Distinct(
            Fonte,
            {"Parâmetro"}
        ),

    AdicionarValue =
        Table.AddColumn(
            Distintos,
            "Value",
            each
                [

                    Valor =
                        [Valor],

                    Tipo =
                        [Tipo],

                    Obrigatório =
                        [Obrigatório],

                    Permitidos =
                        [Permitidos],

                    Origem =
                        [Origem]

                ],
            type record
        ),

    Renomear =
        Table.RenameColumns(
            AdicionarValue,
            {
                {
                    "Parâmetro",
                    "Name"
                }
            }
        ),

    Resultado =
        Table.SelectColumns(
            Renomear,
            {
                "Name",
                "Value"
            }
        )

in

    Resultado


/*
let
    // Normaliza os parâmetros provenientes da tabela do Excel
    ParametrosExcelBase =
        Table.TransformColumns(
            Table.SelectRows(
                srcParametrosExcel,
                each
                    [Parâmetro] <> null
                        and Text.Trim(Text.From([Parâmetro])) <> ""
            ),
            {
                {
                    "Parâmetro",
                    each Text.Trim(Text.From(_)),
                    type text
                }
            }
        ),

    ParametrosExcelComTipo =
        if Table.HasColumns(ParametrosExcelBase, "Tipo") then
            ParametrosExcelBase
        else
            Table.AddColumn(
                ParametrosExcelBase,
                "Tipo",
                each null,
                type nullable text
            ),

    ParametrosExcelComObrigatorio =
        if Table.HasColumns(ParametrosExcelComTipo, "Obrigatório") then
            ParametrosExcelComTipo
        else
            Table.AddColumn(
                ParametrosExcelComTipo,
                "Obrigatório",
                each null,
                type nullable logical
            ),

    ParametrosExcelComPermitidos =
        if Table.HasColumns(ParametrosExcelComObrigatorio, "Permitidos") then
            ParametrosExcelComObrigatorio
        else
            Table.AddColumn(
                ParametrosExcelComObrigatorio,
                "Permitidos",
                each null,
                type nullable text
            ),

    ParametrosExcelTipos =
        Table.TransformColumns(
            ParametrosExcelComPermitidos,
            {
                {
                    "Tipo",
                    each fxParametroTipo(_),
                    type nullable type
                },
                {
                    "Obrigatório",
                    each
                        fxConversor(
                            _,
                            type logical
                        ),
                    type nullable logical
                }
            }
        ),

    ParametrosExcel =
        Table.AddColumn(
            Table.AddColumn(
                ParametrosExcelTipos,
                "Permitidos2",
                each
                    if [Permitidos] = null then
                        null
                    else
                        List.Transform(
                            fxLista([Permitidos]),
                            (v) =>
                                fxConversor(
                                    v,
                                    [Tipo]
                                )
                        ),
                type nullable list
            ),
            "Origem",
            each "Excel",
            type text
        ),

    ParametrosExcelFinal =
        Table.RenameColumns(
            Table.RemoveColumns(
                ParametrosExcel,
                {"Permitidos"}
            ),
            {
                {"Permitidos2", "Permitidos"}
            }
        ),

    ParametrosPowerQuery =
        Table.FromRecords(
            List.Transform(
                Record.FieldNames(cfgParametrosPowerQuery),
                (Nome) =>
                    let
                        Valor =
                            Record.Field(
                                cfgParametrosPowerQuery,
                                Nome
                            ),

                        Metadata =
                            Value.Metadata(Valor)
                    in
                        [
                            Parâmetro = Text.Trim(Nome),
                            Valor = Valor,
                            Tipo =
                                fxParametroTipo(
                                    Record.FieldOrDefault(
                                        Metadata,
                                        "Type"
                                    )
                                ),
                            Obrigatório =
                                Record.FieldOrDefault(
                                    Metadata,
                                    "IsParameterQueryRequired"
                                ),
                            Permitidos = null,
                            Origem = "PowerQuery"
                        ]
            ),
            type table
            [
                Parâmetro = text,
                Valor = any,
                Tipo = nullable type,
                Obrigatório = nullable logical,
                Permitidos = nullable list,
                Origem = text
            ]
        ),

    ParametrosConsolidados =
        Table.Distinct(
            Table.Combine(
                {
                    ParametrosExcelFinal,
                    ParametrosPowerQuery
                }
            ),
            {"Parâmetro"}
        ),

    ComValueRecord = Table.AddColumn(
        ParametrosConsolidados,
        "Value",
        each [
            Valor = [Valor],
            Tipo = [Tipo],
            Obrigatório = [Obrigatório],
            Permitidos = [Permitidos],
            Origem = [Origem]
        ],
        type record
    ),
    RenomeadoName = Table.RenameColumns(ComValueRecord, {{"Parâmetro", "Name"}}),
    ParametrosCatalogo = Table.SelectColumns(RenomeadoName, {"Name", "Value"})

in
    ParametrosCatalogo
*/;

shared fxParametroTipo = (tipo as nullable text) as nullable type =>

let
    Nome =
        if tipo = null then
            null
        else
            Text.Upper(
                Text.Trim(
                    Text.From(tipo)
                )
            )
in
    if Nome = null then

        null

    else if Record.HasFields(
        cfgTiposDados,
        Nome
    ) then

        Record.Field(
            cfgTiposDados,
            Nome
        )

    else

        error Error.Record(
            "Tipo inválido",
            Text.Format(
                "O tipo '#{0}' não está cadastrado em tbParametrosTipos.",
                {tipo}
            ),
            [
                Tipo = tipo,
                TiposDisponiveis =
                    Text.Combine(
                        List.Sort(
                            Record.FieldNames(
                                cfgTiposDados
                            )
                        ),
                        ", "
                    )
            ]
        );

shared cfgTiposDados = /*
Tipos de dados
Define os tipos de parâmetros aceitos pelas funções de tratamento e validação,
determinando como os valores informados nas configurações serão interpretados
e convertidos durante a execução do framework.
*/

[
    #"BOOLEANO" = type logical,
    #"DATA" = type date,
    #"DATA E HORA" = type datetime,
    #"DATA/HORA/FUSO" = type datetimezone,
    #"DURAÇÃO" = type duration,
    #"HORA" = type time,
    #"LISTA" = type list,
    #"NÚMERO DECIMAL" = type number,
    #"NÚMERO INTEIRO" = Int64.Type,
    #"QUALQUER VALOR" = type any,
    #"TEXTO" = type text,
    #"LOGICAL" = type logical,
    #"DATE" = type date,
    #"DATETIME" = type datetime,
    #"DATETIMEZONE" = type datetimezone,
    #"DURATION" = type duration,
    #"TIME" = type time,
    #"LIST" = type list,
    #"NUMBER" = type number,
    #"INT64" = Int64.Type,
    #"ANY" = type any,
    #"TEXT" = type text
];

shared cfgTiposObjetos = /*
------------------------------------------------------------------------------
Consulta.....: cfgTiposObjetos

Descrição....:
Catálogo dos tipos nativos da linguagem M suportados pelo framework.

Centraliza a definição dos tipos utilizados pelas rotinas de inspeção e
identificação de objetos, evitando listas codificadas nas funções.

Esta consulta é consumida principalmente por:

    • fxIdentificarTipoObjeto
    • fxIdentificarObjeto
    • cfgObjetos

Estrutura:

    Kind
        Nome técnico do tipo na linguagem M.

    Nome
        Nome amigável utilizado pelo framework.

    Type
        Tipo M utilizado nas comparações com Type.Is().

    Categoria
        Classificação conceitual do tipo.

    IsStructured
        Indica se o tipo representa uma estrutura composta.

------------------------------------------------------------------------------
*/


    [

        Table = [
            Kind = "Table",
            Nome = "Tabela",
            Type = type table,
            Categoria = "Estruturado",
            IsStructured = true
        ],

        Record = [
            Kind = "Record",
            Nome = "Registro",
            Type = type record,
            Categoria = "Estruturado",
            IsStructured = true
        ],

        List = [
            Kind = "List",
            Nome = "Lista",
            Type = type list,
            Categoria = "Estruturado",
            IsStructured = true
        ],

        Function = [
            Kind = "Function",
            Nome = "Função",
            Type = type function,
            Categoria = "Executável",
            IsStructured = false
        ],

        Action = [
            Kind = "Action",
            Nome = "Ação",
            Type = type action,
            Categoria = "Executável",
            IsStructured = false
        ],

        Text = [
            Kind = "Text",
            Nome = "Texto",
            Type = type text,
            Categoria = "Escalar",
            IsStructured = false
        ],

        Logical = [
            Kind = "Logical",
            Nome = "Lógico",
            Type = type logical,
            Categoria = "Escalar",
            IsStructured = false
        ],

        Int64 = [
            Kind = "Int64",
            Nome = "Inteiro 64 bits",
            Type = Int64.Type,
            Categoria = "Numérico",
            IsStructured = false
        ],

        Number = [
            Kind = "Number",
            Nome = "Número",
            Type = type number,
            Categoria = "Numérico",
            IsStructured = false
        ],

        Date = [
            Kind = "Date",
            Nome = "Data",
            Type = type date,
            Categoria = "Data e Hora",
            IsStructured = false
        ],

        Time = [
            Kind = "Time",
            Nome = "Hora",
            Type = type time,
            Categoria = "Data e Hora",
            IsStructured = false
        ],

        DateTime = [
            Kind = "DateTime",
            Nome = "Data e Hora",
            Type = type datetime,
            Categoria = "Data e Hora",
            IsStructured = false
        ],

        DateTimeZone = [
            Kind = "DateTimeZone",
            Nome = "Data e Hora com Fuso",
            Type = type datetimezone,
            Categoria = "Data e Hora",
            IsStructured = false
        ],

        Duration = [
            Kind = "Duration",
            Nome = "Duração",
            Type = type duration,
            Categoria = "Data e Hora",
            IsStructured = false
        ],

        Binary = [
            Kind = "Binary",
            Nome = "Binário",
            Type = type binary,
            Categoria = "Binário",
            IsStructured = false
        ],

        Type = [
            Kind = "Type",
            Nome = "Tipo",
            Type = type type,
            Categoria = "Especial",
            IsStructured = false
        ],

        Null = [
            Kind = "Null",
            Nome = "Nulo",
            Type = type null,
            Categoria = "Especial",
            IsStructured = false
        ],

        Any = [
            Kind = "Any",
            Nome = "Qualquer",
            Type = type any,
            Categoria = "Especial",
            IsStructured = false
        ]

    ]

;

shared stgParametrosFormatosArquivos = let
    Fonte =
        srcParametrosFormatosArquivos,

    ColunasRenomeadas =
        Table.RenameColumns(
            Fonte,
            {
                {"Formato", "Name"},
                {"Extensão", "Value"}
            },
            MissingField.Ignore
        ),

    LinhasValidas =
        Table.SelectRows(
            ColunasRenomeadas,
            each
                [Name] <> null
                and Text.Trim(Text.From([Name])) <> ""
                and [Value] <> null
                and Text.Trim(Text.From([Value])) <> ""
        ),

    Normalizado =
        Table.TransformColumns(
            LinhasValidas,
            {
                {
                    "Name",
                    each
                        Text.Upper(
                            Text.Trim(
                                Text.From(_)
                            )
                        ),
                    type text
                },
                {
                    "Value",
                    each
                        let
                            Extensao =
                                Text.Lower(
                                    Text.Trim(
                                        Text.From(_)
                                    )
                                )
                        in
                            if Text.StartsWith(
                                Extensao,
                                "."
                            )
                            then
                                Extensao
                            else
                                "." & Extensao,
                    type text
                }
            }
        ),

    DuplicatasRemovidas =
        Table.Distinct(
            Normalizado,
            {
                "Name",
                "Value"
            }
        )

in
    DuplicatasRemovidas;

shared stgParametrosTratamentos = let

    Fonte =
        srcParametrosTratamentos,

    LinhasValidas = Table.SelectRows(
        Fonte,
        each [Ativo] = true and
             [Descrição] <> null and Text.Trim(Text.From([Descrição])) <> "" and
             [Código] <> null and Text.Trim(Text.From([Código])) <> "" and
             [Função] <> null and Text.Trim(Text.From([Função])) <> "" and
             [TipoEntrada] <> null and Text.Trim(Text.From([TipoEntrada])) <> "" and
             [TipoSaída] <> null and Text.Trim(Text.From([TipoSaída])) <> ""
    ),
    
    Tratamentos =
        Table.TransformColumns(
            LinhasValidas,
            {
                {
                    "Descrição",
                    each
                        Text.Upper(
                            Text.Trim(
                                Text.From(_)
                            )
                        )
                ,
                    type text
                },
                {
                    "Código",
                    each
                        Text.Upper(
                            Text.Trim(
                                Text.From(_)
                            )
                        ),
                    type text
                },
                {
                    "Função",
                    each
                        Record.FieldOrDefault(
                            cfgFuncoesPowerQuery,
                            Text.Trim(
                                Text.From(_)
                            ),
                            null
                        ),
                    type function
                },
                {
                    "TipoEntrada",
                    each
                        Text.Trim(
                            Text.From(_)
                        ),
                    type text
                },
                {
                    "TipoSaída",
                    each
                        Text.Trim(
                            Text.From(_)
                        ),
                    type text
                }
            }
        ),

    BaseExcel =
        Table.FromColumns(
            {
                Tratamentos[Descrição],
                Tratamentos[Código],
                Tratamentos[Função],
                Tratamentos[TipoEntrada],
                Tratamentos[TipoSaída],
                Tratamentos[Padrão],
                Tratamentos[Ativo]
            },
            {
                "Name",
                "Código",
                "Função",
                "TipoEntrada",
                "TipoSaída",
                "Padrão",
                "Ativo"
            }
        ),

    BaseCodigo =
        Table.FromColumns(
            {
                Tratamentos[Código],
                Tratamentos[Código],
                Tratamentos[Função],
                Tratamentos[TipoEntrada],
                Tratamentos[TipoSaída],
                Tratamentos[Padrão],
                Tratamentos[Ativo]
            },
            {
                "Name",
                "Código",
                "Função",
                "TipoEntrada",
                "TipoSaída",
                "Padrão",
                "Ativo"
            }
        ),

    TratamentosExpandido =
        Table.Combine(
            {
                BaseExcel,
                BaseCodigo
            }
        ),

    TratamentosUnicos =
        Table.Distinct(
            TratamentosExpandido,
            {"Name"}
        )

in

    TratamentosUnicos;

shared stgParametrosValidacoes = let

    Fonte =
        srcParametrosValidacoes,

    LinhasValidas = Table.SelectRows(
        Fonte,
        each 
             [Ativo] = true and
             [Descrição] <> null and Text.Trim(Text.From([Descrição])) <> "" and
             [Código] <> null and Text.Trim(Text.From([Código])) <> "" and
             [Função] <> null and Text.Trim(Text.From([Função])) <> "" and
             [TipoEntrada] <> null and Text.Trim(Text.From([TipoEntrada])) <> "" and
             [TipoSaída] <> null and Text.Trim(Text.From([TipoSaída])) <> "" and
             [Severidade] <> null and Text.Trim(Text.From([Severidade])) <> ""
    ),

    Validacoes =
        Table.TransformColumns(
            LinhasValidas,
            {
                {
                    "Descrição",
                    each
                        Text.Upper(
                            Text.Trim(
                                Text.From(_)
                            )
                        ),
                    type text
                },
                {
                    "Código",
                    each
                        Text.Upper(
                            Text.Trim(
                                Text.From(_)
                            )
                        ),
                    type text
                },
                {
                    "Função",
                    each
                        Record.FieldOrDefault(
                            cfgFuncoesPowerQuery,
                            Text.Trim(
                                Text.From(_)
                            ),
                            null
                        ),
                    type function
                },
                {
                    "Severidade",
                    each
                        Text.Upper(
                            Text.Trim(
                                Text.From(_)
                            )
                        ),
                    type text
                },
                {
                    "TipoEntrada",
                    each
                        Text.Trim(
                            Text.From(_)
                        ),
                    type text
                },
                {
                    "TipoSaída",
                    each
                        Text.Trim(
                             Text.From(_)
                        ),
                    type text
                }
            }
        ),

    ValidacoesCompleto =
        Table.InsertRows(
            Validacoes,
            0,
            {
                [
                    Descrição = "OBRIGATÓRIO",
                    Código = "REQUIRED",
                    Função = fxSchemaValidacaoREQUIRED,
                    Severidade = "ERROR",
                    TipoEntrada = "Qualquer valor",
                    TipoSaída = "Qualquer valor",
                    Padrão = true,
                    Ativo = true
                ]
            }
        ),

    BaseExcel =
        Table.FromColumns(
            {
                ValidacoesCompleto[Descrição],
                ValidacoesCompleto[Código],
                ValidacoesCompleto[Função],
                ValidacoesCompleto[Severidade],
                ValidacoesCompleto[TipoEntrada],
                ValidacoesCompleto[TipoSaída],
                ValidacoesCompleto[Padrão],
                ValidacoesCompleto[Ativo]
            },
            {
                "Name",
                "Código",
                "Função",
                "Severidade",
                "TipoEntrada",
                "TipoSaída",
                "Padrão",
                "Ativo"
            }
        ),

    BaseCodigo =
        Table.FromColumns(
            {
                ValidacoesCompleto[Código],
                ValidacoesCompleto[Código],
                ValidacoesCompleto[Função],
                ValidacoesCompleto[Severidade],
                ValidacoesCompleto[TipoEntrada],
                ValidacoesCompleto[TipoSaída],
                ValidacoesCompleto[Padrão],
                ValidacoesCompleto[Ativo]
            },
            {
                "Name",
                "Código",
                "Função",
                "Severidade",
                "TipoEntrada",
                "TipoSaída",
                "Padrão",
                "Ativo"
            }
        ),

    ValidacoesExpandido =
        Table.Combine(
            {
                BaseExcel,
                BaseCodigo
            }
        ),

    ValidacoesUnicos =
        Table.Distinct(
            ValidacoesExpandido,
            {"Name"}
        )

in

    ValidacoesUnicos;

shared stgParametrosSeveridades = let

    Fonte =
        srcParametrosSeveridades,

    LinhasValidas = Table.SelectRows(
        Fonte,
        each [Descrição] <> null and Text.Trim(Text.From([Descrição])) <> "" and
             [Código] <> null and Text.Trim(Text.From([Código])) <> ""
    ),

    Severidades =
        Table.TransformColumns(
            LinhasValidas,
            {
                {
                    "Código",
                    each
                        Text.Upper(
                            Text.Trim(
                                Text.From(_)
                            )
                        ),
                    type text
                },
                {
                    "Descrição",
                    each
                        Text.Upper(
                            Text.Trim(
                                Text.From(_)
                            )
                        ),
                    type text
                },
                {
                    "Ordem",
                    each
                        if _ = null then
                            null
                        else
                            Int64.From(_),
                    Int64.Type
                },
                {
                    "Bloqueia",
                    each
                        if _ = null then
                            false
                        else
                            Logical.From(_),
                    type logical
                }
            }
        ),

    Distintos =
        Table.Distinct(
            Severidades,
            {"Código"}
        )

in

    Distintos;

shared stgSchema = let

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

    Fonte =
        srcSchema,

//--------------------------------------------------------------------------
// Linhas válidas
//--------------------------------------------------------------------------

    LinhasValidas =

        Table.SelectRows(

            Fonte,

            each

                [Ativo] = true
                and [Tabela] <> null
                and Text.Trim(Text.From([Tabela])) <> ""
                and [Coluna] <> null
                and Text.Trim(Text.From([Coluna])) <> ""

        ),

//--------------------------------------------------------------------------
// Normalização
//--------------------------------------------------------------------------

    Normalizado =

        Table.TransformColumns(

            LinhasValidas,

            {

                {
                    "Tabela",
                    each Text.Trim(Text.From(_)),
                    type text
                },

                {
                    "Coluna",
                    each Text.Trim(Text.From(_)),
                    type text
                },

                {
                    "Tipo",
                    each fxParametroTipo(_),
                    type type
                },

                {
                    "Obrigatório",
                    each fxConversor(_, type logical),
                    type logical
                },

                {
                    "Ativo",
                    each fxConversor(_, type logical),
                    type logical
                },

                {
                    "Ordem",

                    each

                        if _ = null
                        or Text.Trim(Text.From(_)) = ""

                        then
                            null

                        else
                            Int64.From(_),

                    Int64.Type

                },

                {
                    "Tratamentos",

                    each

                        let

                            Lista =

                                if _ = null then

                                    {}

                                else

                                    List.Distinct(

                                        List.RemoveNulls(

                                            List.Transform(

                                                fxLista(_),

                                                each

                                                    let

                                                        Texto =
                                                            Text.Trim(
                                                                Text.From(_)
                                                            )

                                                    in

                                                        if Texto = "" then
                                                            null
                                                        else
                                                            Texto

                                            )

                                        )

                                    )

                        in

                            if List.IsEmpty(Lista) then
                                null
                            else
                                List.Buffer(Lista),

                    type nullable list

                },

                {
                    "Validações",

                    each

                        let

                            Lista =

                                if _ = null then

                                    {}

                                else

                                    List.Distinct(

                                        List.RemoveNulls(

                                            List.Transform(

                                                fxLista(_),

                                                each

                                                    let

                                                        Texto =
                                                            Text.Trim(
                                                                Text.From(_)
                                                            )

                                                    in

                                                        if Texto = "" then
                                                            null
                                                        else
                                                            Texto

                                            )

                                        )

                                    )

                        in

                            if List.IsEmpty(Lista) then
                                null
                            else
                                List.Buffer(Lista),

                    type nullable list

                }

            }

        ),

//--------------------------------------------------------------------------
// Duplicatas
//--------------------------------------------------------------------------

    DuplicatasRemovidas =

        Table.Distinct(

            Normalizado,

            {

                "Tabela",
                "Coluna"

            }

        ),

//--------------------------------------------------------------------------
// Ordenação
//--------------------------------------------------------------------------

    Resultado =

        Table.ReorderColumns(

            DuplicatasRemovidas,

            {

                "Tabela",
                "Coluna",
                "Tipo",
                "Obrigatório",
                "Ordem",
                "Tratamentos",
                "Validações",
                "Ativo"

            }

        )

in

    Resultado



/*
let

    Fonte =
        srcSchema,

    LinhasValidas =
        Table.SelectRows(
            Fonte,
            each
                [Ativo] = true
                and [Tabela] <> null
                and Text.Trim(Text.From([Tabela])) <> ""
                and [Coluna] <> null
                and Text.Trim(Text.From([Coluna])) <> ""
        ),

    ColunasNormalizadas =
        Table.TransformColumns(
            LinhasValidas,
            {
                {
                    "Tabela",
                    each Text.Trim(Text.From(_)),
                    type text
                },
                {
                    "Coluna",
                    each Text.Trim(Text.From(_)),
                    type text
                },
                {
                    "Tipo",
                    each fxParametroTipo(_),
                    type type
                },
                {
                    "Obrigatório",
                    each fxConversor(_, type logical),
                    type logical
                },
                {
                    "Ativo",
                    each fxConversor(_, type logical),
                    type logical
                },
                {
                    "Ordem",
                    each
                        if _ = null or Text.Trim(Text.From(_)) = "" then
                            null
                        else
                            Int64.From(_),
                    Int64.Type
                },
                {
                    "Tratamentos",

                    each

                        let

                            Regras =
                                if _ = null then
                                    {}
                                else
                                    List.Distinct(
                                        fxLista(_)
                                    ),

                            TratamentosInterpretados =
                                List.Transform(
                                    Regras,
                                    (Texto) =>

                                        let

                                            RegraInterpretada =
                                                fxSchemaInterpretarTratamentos(
                                                    Texto
                                                ),

                                            Codigo =
                                                Text.Upper(
                                                    Text.Trim(
                                                        RegraInterpretada[Nome]
                                                    )
                                                ),

                                            Definicao =
                                                Record.FieldOrDefault(
                                                    cfgParametrosTratamentos,
                                                    Codigo
                                                )

                                        in

                                            if Definicao = null then

                                                error Error.Record(
                                                    "Tratamento inválido",
                                                    "Tratamento não cadastrado ou foi inativado.",
                                                    [
                                                        Código = Codigo
                                                    ]
                                                )

                                            else

                                                [
                                                    Definição = Definicao,
                                                    Regra = Texto,
                                                    Parametros =
                                                        if List.IsEmpty(RegraInterpretada[Parametros]) then
                                                            null
                                                        else
                                                            List.Buffer(RegraInterpretada[Parametros])
                                                ]
                                )

                        in

                            if List.IsEmpty(TratamentosInterpretados) then
                                null
                            else
                                List.Buffer(TratamentosInterpretados),

                    type nullable list
                }
            }
        ),

    Validacoes =
        Table.AddColumn(
            ColunasNormalizadas,
            "Validações2",

            each

                let

                    Regras =
                        if [Validações] = null then
                            {}
                        else
                            List.Distinct(
                                fxLista([Validações])
                            ),

                    ValidacoesInterpretadas =
                        List.Transform(
                            Regras,
                            (Texto) =>

                                let

                                    RegraInterpretada =
                                        fxSchemaInterpretarValidacao(
                                            Texto
                                        ),

                                    Codigo =
                                        Text.Upper(
                                            Text.Trim(
                                                RegraInterpretada[Nome]
                                            )
                                        ),

                                    Definicao =
                                        Record.FieldOrDefault(
                                            cfgParametrosValidacoes,
                                            Codigo
                                        )

                                in

                                    if Definicao = null then

                                        error Error.Record(
                                            "Validação inválida",
                                            "Validação não cadastrada.",
                                            [
                                                Código = Codigo
                                            ]
                                        )

                                    else

                                        [
                                            Definição = Definicao,
                                            Regra = Texto,
                                            Parametros =
                                                if List.IsEmpty(RegraInterpretada[Parametros]) then
                                                    null
                                                else
                                                    List.Buffer(RegraInterpretada[Parametros])
                                        ]
                        ),

                    Required =
                        if [Obrigatório] then
                            {
                                [
                                    Definição = cfgParametrosValidacoes[REQUIRED],
                                    Regra = "REQUIRED",
                                    Parametros = null
                                ]
                            }
                        else
                            {},

                    Resultado =
                        List.Combine(
                            {
                                Required,
                                ValidacoesInterpretadas
                            }
                        )

                in

                    if List.IsEmpty(Resultado) then
                        null
                    else
                        List.Buffer(Resultado),

            type nullable list
        ),

    ColunasAjustadas =
        Table.RenameColumns(
            Table.RemoveColumns(
                Validacoes,
                {"Validações"}
            ),
            {
                {"Validações2", "Validações"}
            }
        ),

    DuplicatasRemovidas =
        Table.Distinct(
            ColunasAjustadas,
            {
                "Tabela",
                "Coluna"
            }
        ),
    ColunasReordenadas = Table.ReorderColumns(DuplicatasRemovidas,{"Tabela", "Coluna", "Tipo", "Obrigatório", "Ordem", "Tratamentos", "Validações", "Ativo"})
in
    ColunasReordenadas
*/;

shared cfgTiposBooleanos = /*
Valores Booleanos
Define os valores reconhecidos pelo framework como equivalentes aos valores
lógicos VERDADEIRO e FALSO, permitindo interpretar diferentes representações
textuais durante o processamento dos dados.
*/

[
    #"TRUE" = true,
    #"FALSE" = false,
    #"SIM" = true,
    #"NÃO" = false,
    #"NAO" = false,
    #"1" = true,
    #"0" = false
];

shared cfgConversores = [
    Any = (v as any, local as text) as any => v,
    Text = (v as any, local as text) as any => Text.From(v, local),
    List = (v as any, local as text) as any => fxLista(v),
    Int64 = (v as any, local as text) as any => Int64.From(v, local),
    Number = (v as any, local as text) as any => Number.From(v, local),
    Date = (v as any, local as text) as any => Date.From(v, local),
    DateTime = (v as any, local as text) as any => DateTime.From(v, local),
    DateTimeZone = (v as any, local as text) as any => DateTimeZone.From(v, local),
    Time = (v as any, local as text) as any => Time.From(v, local),
    Duration = (v as any, local as text) as any => Duration.From(v),
    Logical = (v as any, local as text) as any => fxBooleano(v, local)
]
;

shared fxLista = (
    valor as any,
    optional separador as nullable text,
    optional removerVazios as nullable logical,
    optional trim as nullable logical
)
as nullable list =>

let
    Separador =
        if separador = null then
            ";"
        else
            separador,

    RemoverVazios =
        if removerVazios = null then
            true
        else
            removerVazios,

    Trim =
        if trim = null then
            true
        else
            trim,

    Lista =
        if valor = null then

            null

        else if Value.Is(valor, type list) then

            valor

        else

            Text.Split(
                Text.From(valor),
                Separador
            ),

    ListaTratada =
        if Lista = null then
            null
        else
            List.Transform(
                Lista,
                each
                    if Trim then
                        Text.Trim(Text.From(_))
                    else
                        Text.From(_)
            ),

    Resultado =
        if ListaTratada = null then
            null
        else if RemoverVazios then
            List.RemoveItems(
                ListaTratada,
                {""}
            )
        else
            ListaTratada

in
    if Resultado = null then
        null
    else
        List.Buffer(
            List.Distinct(Resultado)
        );

shared fxBooleano = (valor as any, optional local as nullable text) as nullable logical =>

let
    Local =
        if local = null then
            "pt-BR"
        else
            local,

    Nome =
        if valor = null then
            null
        else
            Text.Upper(
                Text.Trim(
                    Text.From(
                        valor,
                        Local
                    )
                )
            )
in
    if Nome = null then

        null

    else if Record.HasFields(
        cfgTiposBooleanos,
        Nome
    ) then

        Record.Field(
            cfgTiposBooleanos,
            Nome
        )

    else

        error Error.Record(
            "Valor lógico inválido",
            Text.Format(
                "O valor '#{0}' não é uma representação válida de lógico.",
                {valor}
            ),
            [
                Valor = valor,
                ValoresPermitidos =
                    Text.Combine(
                        List.Sort(
                            Record.FieldNames(cfgTiposBooleanos)
                        ),
                        ", "
                    )
            ]
        );

shared fxComoTabela = (Valor as any) as table =>

let
    Resultado =
        if Valor = null then

            #table(
                type table
                [
                    Value = any
                ],
                {}
            )

        else if Value.Is(
            Valor,
            type table
        ) then

            Valor

        else if Value.Is(
            Valor,
            type record
        ) then

            Record.ToTable(
                Valor
            )

        else if Value.Is(
            Valor,
            type list
        ) then

            Table.FromList(
                Valor,
                Splitter.SplitByNothing(),
                {"Value"}
            )

        else

            #table(
                type table
                [
                    Value = any
                ],
                {
                    {Valor}
                }
            )

in
    Resultado;

shared fxConversor = (
    valor as any,
    tipo as nullable type,
    optional local as nullable text
) as any =>

let
    Local = if local = null then "pt-BR" else local,
    TipoDestino = if tipo = null then type any else tipo,

    Chave =
        if TipoDestino = type any then "Any"
        else if TipoDestino = type text then "Text"
        else if TipoDestino = type list then "List"
        else if TipoDestino = Int64.Type then "Int64"
        else if TipoDestino = type number then "Number"
        else if TipoDestino = type date then "Date"
        else if TipoDestino = type datetime then "DateTime"
        else if TipoDestino = type datetimezone then "DateTimeZone"
        else if TipoDestino = type time then "Time"
        else if TipoDestino = type duration then "Duration"
        else if TipoDestino = type logical then "Logical"
        else error Error.Record(
            "Tipo não suportado",
            "A função não possui conversor para o tipo solicitado.",
            [Tipo = TipoDestino]
        ),

    Conversor = Record.Field(cfgConversores, Chave),
    Resultado = Conversor(valor, Local)
in
    Resultado;

shared fxConectorLocal = (Caminho as text) as table =>

Folder.Files(
    Caminho
);

shared fxConectorOracle = (
    Host as text,
    Banco as nullable text,
    SQL as text
)
as table =>

let
    Resultado =
        Oracle.Database(
            Host,
            [
                Query = SQL
            ]
        )

in
    Resultado;

shared fxConectorSQLServer = (
    Host as text,
    Banco as text,
    SQL as text
)
as table =>

let
    Resultado =
        Sql.Database(
            Host,
            Banco,
            [
                Query = SQL
            ]
        )

in
    Resultado;

shared fxConectorPostgreSQL = (
    Host as text,
    Banco as text,
    SQL as text
)
as table =>

let
    Resultado =
        PostgreSQL.Database(
            Host,
            Banco,
            [
                Query = SQL
            ]
        )

in
    Resultado;

shared fxConectorMySQL = (
    Host as text,
    Banco as text,
    SQL as text
)
as table =>

let
    Resultado =
        MySQL.Database(
            Host,
            Banco,
            [
                Query = SQL
            ]
        )

in
    Resultado;

shared fxConectorSharePoint = (Site as text, Caminho as text) as table =>

let
    Partes =
        List.Select(
            Text.Split(
                Text.Replace(Caminho, "\", "/"),
                "/"
            ),
            each _ <> ""
        ),

    Biblioteca =
        fxParametro("Biblioteca_SharePoint", "Documents"),

    Fonte =
        SharePoint.Contents(Site),

    BibliotecaRaiz =
        try
            Fonte{[Name = Biblioteca]}[Content]
        otherwise
            error Error.Record(
                "Biblioteca inexistente",
                Text.Format(
                    "A biblioteca '#{0}' não foi encontrada no site '#{1}'.",
                    {
                        Biblioteca,
                        Site
                    }
                ),
                [
                    Biblioteca = Biblioteca,
                    Site = Site
                ]
            ),

    PastaRaiz =
        List.Accumulate(
            Partes,
            BibliotecaRaiz,
            (Estado, PastaAtual) =>
                try
                    Estado{[Name = PastaAtual]}[Content]
                otherwise
                    error Error.Record(
                        "Pasta inexistente",
                        Text.Format(
                            "A pasta '#{0}' não foi encontrada.",
                            {PastaAtual}
                        ),
                        [
                            Pasta = PastaAtual
                        ]
                    )
        ),

    ArquivosRecursivos =
        (Tabela as table) as table =>

        let
            Arquivos =
                Table.SelectRows(
                    Tabela,
                    each
                        try Value.Is([Content], Binary.Type)
                        otherwise false
                ),

            Pastas =
                Table.SelectRows(
                    Tabela,
                    each
                        try Value.Is([Content], Table.Type)
                        otherwise false
                ),

            ArquivosFilhos =
                List.RemoveNulls(
                    List.Transform(
                        Pastas[Content],
                        each
                            try
                                @ArquivosRecursivos(_)
                            otherwise
                                null
                    )
                ),

            Resultado =
                if List.Count(ArquivosFilhos) > 0 then
                    Table.Combine(
                        List.Combine(
                            {
                                {Arquivos},
                                ArquivosFilhos
                            }
                        )
                    )
                else
                    Arquivos

        in
            Resultado,

    Resultado =
        ArquivosRecursivos(
            PastaRaiz
        )

in
    Resultado;

[ Description = "Consultas existentes do PowerQuery" ]
shared diagConsultasPQ = let

//--------------------------------------------------------------------------
// Objetos do Power Query
//--------------------------------------------------------------------------

    Objetos =

        Record.FieldValues(

            cfgObjetos[PowerQuery]

        ),

//--------------------------------------------------------------------------
// Tabela
//--------------------------------------------------------------------------

    Consultas =

        Table.FromRecords(

            List.Transform(

                Objetos,

                each

                    [

                        Consulta = [Nome],

                        Categoria = [Categoria],

                        Tipo = [Tipo]

                    ]

            )

        ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        Table.Sort(

            Consultas,

            {

                {"Categoria", Order.Ascending},

                {"Consulta", Order.Ascending}

            }

        )

in

    Resultado;

shared diagTabelasExcel = let

    Fonte =

        Record.ToTable(
            cfgTabelasExcel
        ),

    Expandir =

        Table.ExpandRecordColumn(

            Fonte,

            "Value",

            {
                "Columns"
            },

            {
                "Colunas"
            }

        ),

    ColunasTexto =

        Table.TransformColumns(

            Expandir,

            {
                {
                    "Colunas",
                    each Text.Combine(_, ";"),
                    type text
                }
            }

        ),

    Renomeadas =

        Table.RenameColumns(

            ColunasTexto,

            {
                {"Name", "Tabela"}
            }

        ),

    Tipado =

        Table.TransformColumnTypes(

            Renomeadas,

            {
                {"Tabela", type text},
                {"Colunas", type text}
            }

        ),

    Resultado =

        Table.SelectColumns(

            Tipado,

            {
                "Tabela",
                "Colunas"
            }

        )

in

    Resultado;

shared stgClientes = let
    Fonte =
        srcClientes,

    Stage =
        fxStgAplicar(
            Fonte,
            parTabelaClientes
        ),

    Resultado =
        Table.Distinct(
            Stage
        )
in
    Resultado;

shared stgProdutos = let
    Fonte =
        srcProdutos,

    Stage =
        fxStgAplicar(
            Fonte,
            parTabelaProdutos
        ),

    Resultado =
        Table.Distinct(
            Stage
        )
in
    Resultado;

shared stgVendas = let
    Fonte =
        srcVendas,

    Preparado =
        fxStgAplicar(
            Fonte,
            parTabelaVendas
        ),

    RegistrosUnicos =
        Table.Distinct(
            Preparado
        )
in
    RegistrosUnicos;

shared nrmClientes = let

    Fonte =
        stgClientes,

    Normalizacao =
        fxNrmAplicar(
            Fonte,
            "tbClientes"
        )

in

    Normalizacao;

shared dimClientes = let
    // Obtém a entidade normalizada.
    Fonte =
        nrmClientes,

    // Seleciona apenas os atributos da dimensão.
    Atributos =
        Fonte,

    // Remove registros duplicados.
    RegistrosUnicos =
        Table.Distinct(
            Atributos
        ),

    // Ordena os registros para garantir estabilidade da chave substituta.
    RegistrosOrdenados =
        Table.Sort(
            RegistrosUnicos,
            {
                {"CPF", Order.Ascending}
            }
        ),

    // Gera a chave substituta da dimensão.
    Chaves =
        Table.AddIndexColumn(
            RegistrosOrdenados,
            "IDCliente",
            1,
            1,
            Int64.Type
        ),

    // Reorganiza as colunas da dimensão.
    ColunasReordenadas =
        Table.ReorderColumns(
            Chaves,
            {
                "IDCliente",
                "CPF",
                "Nome",
                "DataNascimento",
                "Cidade",
                "Estado"
            }
        ),
    
    Buffer = Table.Buffer(ColunasReordenadas)
in
    Buffer;

shared fatoVendas = let
    // Obtém os dados normalizados.
    Fonte =
        nrmVendas,
    // Resolve a chave da dimensão Cliente.
    Cliente =
        Table.NestedJoin(
            Fonte,
            {"CPF"},
            dimClientes,
            {"CPF"},
            "_Cliente",
            JoinKind.Inner
        ),
    ClienteExpandido =
        Table.ExpandTableColumn(
            Cliente,
            "_Cliente",
            {"IDCliente"}
        ),
    // Resolve a chave da dimensão Produto.
    Produto =
        Table.NestedJoin(
            ClienteExpandido,
            {"CódigoProduto"},
            dimProdutos,
            {"Código"},
            "_Produto",
            JoinKind.Inner
        ),
    ProdutoExpandido =
        Table.ExpandTableColumn(
            Produto,
            "_Produto",
            {"IDProduto"}
        ),
    // Seleciona apenas as colunas do modelo dimensional.
    Colunas =
        Table.SelectColumns(
            ProdutoExpandido,
            {
                "Data",
                "IDCliente",
                "IDProduto",
                "Quantidade",
                "ValorUnitário",
                "ValorTotal"
            }
        ),
    // Organiza as colunas da tabela fato.
    ColunasReordenadas =
        Table.ReorderColumns(
            Colunas,
            {
                "Data",
                "IDCliente",
                "IDProduto",
                "Quantidade",
                "ValorUnitário",
                "ValorTotal"
            }
        )
in
    ColunasReordenadas;

shared fxCaminhoLocal = (Chave as text) as text =>

let
    Valor =
        Text.Trim(
            fxParametro(Chave)
        ),

    Corrigido =
        Text.Replace(
            Valor,
            "/",
            "\"
        ),

    SemBarraFinal =
        Text.TrimEnd(
            Corrigido,
            {"\"}
        )

in
    SemBarraFinal;

shared fxCaminhoSharePoint = (Chave as text) as text =>

let
    ValorOriginal =
        Text.Trim(
            fxParametro(Chave)
        ),

    SeparadoresCorrigidos =
        Text.Replace(
            ValorOriginal,
            "\",
            "/"
        ),

    ComBarraInicial =
        if Text.StartsWith(
            SeparadoresCorrigidos,
            "/"
        )
        then
            SeparadoresCorrigidos
        else
            "/" & SeparadoresCorrigidos,

    SemBarraFinal =
        Text.TrimEnd(
            ComBarraInicial,
            {"/"}
        )

in
    SemBarraFinal;

shared fxODataRequest = (
    optional Url as nullable text
)
as any =>

let
    Endpoint =
        if Url = null then
            fxParametro("Endpoint_REST_Base")
                & fxParametro("Endpoint_REST_Path", "")
        else
            Url,

    Resultado =
        OData.Feed(
            Endpoint
        )

in
    Resultado;

shared fxSOAPRequest = (
    optional Envelope as nullable text,
    optional Headers as nullable record
)
as binary =>

let
    BaseUrl =
        fxParametro("Endpoint_REST_Base"),

    Path =
        fxParametro("Endpoint_REST_Path", ""),

    Timeout =
        fxParametro("Timeout_REST", 30),

    Tentativas =
        fxParametro("Tentativas_REST", 3),

    Corpo =
        if Envelope = null then

            error Error.Record(
                "Envelope SOAP inexistente",
                "Nenhum envelope SOAP foi informado para a requisição.",
                [
                    Endpoint = BaseUrl,
                    RelativePath = Path
                ]
            )

        else

            Envelope,

    Cabecalhos =
        Record.Combine(
            {
                [
                    #"Content-Type" = "text/xml; charset=utf-8"
                ],
                fxRESTHeaders(),
                if Headers = null then [] else Headers
            }
        ),

    StatusReprocessar =
        {
            408,
            429,
            500,
            502,
            503,
            504
        },

    Requisitar =
        (Tentativa as number) as binary =>

        let
            Resposta =
                try
                    Web.Contents(
                        BaseUrl,
                        [
                            RelativePath = Path,
                            Headers = Cabecalhos,
                            Content = Text.ToBinary(Corpo, TextEncoding.Utf8),
                            Timeout = #duration(0,0,0,Timeout),
                            IsRetry = Tentativa > 0,
                            ManualStatusHandling = StatusReprocessar
                        ]
                    ),

            Binario =
                if Resposta[HasError] then
                    null
                else
                    Resposta[Value],

            Status =
                if Resposta[HasError] then
                    null
                else
                    try
                        Value.Metadata(Binario)[Response.Status]
                    otherwise
                        200,

            DeveRepetir =
                (Resposta[HasError]
                    or List.Contains(StatusReprocessar, Status))
                and Tentativa < Tentativas,

            Resultado =
                if DeveRepetir then

                    Function.InvokeAfter(
                        () => @Requisitar(Tentativa + 1),
                        #duration(
                            0,
                            0,
                            0,
                            Number.Power(2, Tentativa)
                        )
                    )

                else if Resposta[HasError] then

                    error Resposta[Error]

                else if List.Contains(StatusReprocessar, Status) then

                    error Error.Record(
                        "Falha HTTP",
                        "A requisição SOAP falhou após todas as tentativas.",
                        [
                            Status = Status,
                            Endpoint = BaseUrl,
                            RelativePath = Path
                        ]
                    )

                else

                    Binario

        in
            Resultado,

    Resultado =
        Requisitar(0)

in
    Resultado;

shared fxRESTRequest = (
    optional RelativePath as nullable text,
    optional Headers as nullable record
)
as binary =>

let
    BaseUrl =
        fxParametro("Endpoint_REST_Base"),

    Path =
        if RelativePath = null then
            fxParametro("Endpoint_REST_Path")
        else
            RelativePath,

    Timeout =
        fxParametro("Timeout_REST", 30),

    Tentativas =
        fxParametro("Tentativas_REST", 3),

    Cabecalhos =
        if Headers = null then
            fxRESTHeaders()
        else
            Record.Combine(
                {
                    fxRESTHeaders(),
                    Headers
                }
            ),

    StatusReprocessar =
        {
            408,
            429,
            500,
            502,
            503,
            504
        },

    Requisitar =
        (Tentativa as number) as binary =>

        let
            Resposta =
                try
                    Web.Contents(
                        BaseUrl,
                        [
                            RelativePath = Path,
                            Headers = Cabecalhos,
                            Timeout = #duration(0, 0, 0, Timeout),
                            IsRetry = Tentativa > 0,
                            ManualStatusHandling = StatusReprocessar
                        ]
                    ),

            Binario =
                if Resposta[HasError] then
                    null
                else
                    Resposta[Value],

            Status =
                if Resposta[HasError] then
                    null
                else
                    try
                        Value.Metadata(Binario)[Response.Status]
                    otherwise
                        200,

            DeveRepetir =
                (Resposta[HasError]
                    or List.Contains(StatusReprocessar, Status))
                and Tentativa < Tentativas,

            Resultado =
                if DeveRepetir then

                    Function.InvokeAfter(
                        () => @Requisitar(Tentativa + 1),
                        #duration(
                            0,
                            0,
                            0,
                            Number.Power(2, Tentativa)
                        )
                    )

                else if Resposta[HasError] then

                    error Resposta[Error]

                else if List.Contains(StatusReprocessar, Status) then

                    error Error.Record(
                        "Falha HTTP",
                        "A requisição REST falhou após todas as tentativas.",
                        [
                            Status = Status,
                            Endpoint = BaseUrl,
                            RelativePath = Path
                        ]
                    )

                else

                    Binario

        in
            Resultado,

    Resultado =
        Requisitar(0)

in
    Resultado;

shared fxRESTHeaders = () as record =>

let
    Resultado =
        cfgRESTHeaders
in
    Resultado;

shared fxOrigemRESTConteudo = () as any =>

let
    Protocolo =
        Text.Upper(
            fxParametro("Protocolo_REST")
        ),

    Resultado =
        if Protocolo = "REST" then

            fxRESTRequest()

        else if Protocolo = "ODATA" then

            fxODataRequest()

        else if Protocolo = "SOAP" then

            // fxSOAPRequest()
            error Error.Record(
                "SOAP requer implementação específica",
                "As consultas SOAP exigem um envelope XML. Utilize diretamente fxSOAPRequest(Envelope)."
            )

        else

            error Error.Record(
                "Protocolo REST não suportado",
                Text.Format(
                    "O protocolo '#{0}' não possui um conector implementado.",
                    {
                        Protocolo
                    }
                ),
                [
                    Protocolo = Protocolo,
                    Permitidos =
                    {
                        "REST",
                        "ODATA",
                        "SOAP"
                    }
                ]
            )

in
    Resultado;

shared fxOrigem = () as any =>

let
    CarregarDadosExternos =
        fxParametro(
            "Carregar_Dados_Externos",
            true
        ),

    FonteDados =
        Text.Upper(
            fxParametro("Fonte_Dados")
        ),

    Resultado =
        if not CarregarDadosExternos then

            #table({}, {})

        else if FonteDados = "ARQUIVOS" then

            srcArquivos

        else if FonteDados = "REST" then

            srcREST

        else if FonteDados = "SGBD" then

            srcSGBD

        else

            error Error.Record(
                "Fonte de dados inválida",
                Text.Format(
                    "A fonte '#{0}' não é suportada.",
                    {
                        FonteDados
                    }
                ),
                [
                    Fonte = FonteDados,
                    Permitidas =
                    {
                        "ARQUIVOS",
                        "REST",
                        "SGBD"
                    }
                ]
            )

in
    Resultado;

shared fxOrigemArquivos = () as table =>

let
    Origem =
        Text.Upper(
            fxParametro("Origem_Arquivos")
        ),
        
    Resultado =
        if Origem = "LOCAL" then

            fxConectorLocal(
                fxCaminhoLocal("Pasta_Local")
            )

        else if Origem = "REMOTA" then

            fxConectorLocal(
                fxCaminhoLocal("Pasta_Remota")
            )

        else if Origem = "SHAREPOINT" then

            fxConectorSharePoint(
                fxParametro("Site_Sharepoint"),
                fxCaminhoSharePoint("Pasta_Sharepoint")
            )

        else

            error Error.Record(
                "Origem inválida",
                "Não existe conector para a origem informada.",
                [
                    Origem = Origem
                ]
            )

in
    Resultado;

shared fxOrigemSGBD = () as table =>

let
    TipoSGBD =
        Text.Upper(
            fxParametro("Tipo_SGBD")
        ),

    Host =
        fxParametro("Host_SGBD"),

    Banco =
        fxParametro("Banco_SGBD"),

    SQL =
        fxParametro("SQL_SGBD"),

    Resultado =
        if TipoSGBD = "ORACLE" then

            fxConectorOracle(
                Host,
                Banco,
                SQL
            )

        else if TipoSGBD = "SQLSERVER" then

            fxConectorSQLServer(
                Host,
                Banco,
                SQL
            )

        else if TipoSGBD = "POSTGRESQL" then

            fxConectorPostgreSQL(
                Host,
                Banco,
                SQL
            )

        else if TipoSGBD = "MYSQL" then

            fxConectorMySQL(
                Host,
                Banco,
                SQL
            )

        else

            error Error.Record(
                "SGBD não suportado",
                Text.Format(
                    "O tipo de SGBD '#{0}' não possui um conector implementado.",
                    {
                        TipoSGBD
                    }
                ),
                [
                    Tipo = TipoSGBD,
                    Permitidos =
                    {
                        "ORACLE",
                        "SQLSERVER",
                        "POSTGRESQL",
                        "MYSQL"
                    }
                ]
            )

in
    Resultado;

shared fxOrigemREST = () as any =>

let
    Conteudo =
        fxOrigemRESTConteudo(),

    Resultado =
        if Value.Is(
            Conteudo,
            Binary.Type
        ) then

            let
                Formato =
                    Text.Upper(
                        fxParametro("Formato_REST")
                    )
            in
                if Formato = "JSON" then

                    Json.Document(
                        Conteudo
                    )

                else if Formato = "XML" then

                    Xml.Tables(
                        Conteudo
                    )

                else

                    error Error.Record(
                        "Formato não suportado",
                        Text.Format(
                            "O formato '#{0}' não possui um interpretador implementado.",
                            {
                                Formato
                            }
                        ),
                        [
                            Formato = Formato,
                            Permitidos =
                            {
                                "JSON",
                                "XML"
                            }
                        ]
                    )

        else

            Conteudo

in
    Resultado;

shared cfgRESTHeaders = /*
Headers REST
Define os cabeçalhos HTTP utilizados nas requisições a serviços que retornam dados
em formato JSON, permitindo configurar informações como autenticação, tipo de
conteúdo e demais parâmetros exigidos pela API.
*/
[
        Accept = "application/json",
        Authorization = "Bearer eyJ...",
        #"x-api-key" = "123456789",
        #"User-Agent" = "Power Query",
        #"Content-Type" = "application/json"
];

shared fxLeitorCSV = (Content as binary) as table =>

let
    Resultado =
        Csv.Document(
            Content,
            [
                Delimiter = ",",
                Encoding = 65001,
                QuoteStyle = QuoteStyle.Csv
            ]
        )

in
    Resultado;

shared fxLeitorExcel = (Content as binary) as table =>

let
    Resultado =
        Excel.Workbook(
            Content,
            null,
            true
        )

in
    Resultado;

shared fxLeitorJSON = (Content as binary) as any =>

let
    Resultado =
        Json.Document(
            Content
        )

in
    Resultado;

shared fxLeitorXML = (Content as binary) as table =>

let
    Resultado =
        Xml.Tables(
            Content
        )

in
    Resultado;

shared fxLeitorPDF = (Content as binary) as table =>

let
    Resultado =
        Pdf.Tables(
            Content
        )

in
    Resultado;

shared fxLeitorArquivo = (Content as binary) as any =>

let
    Formato =
        Text.Upper(
            fxParametro("Formato_Arquivo")
        ),

    Resultado =
        if Formato = "CSV" then

            fxLeitorCSV(
                Content
            )

        else if Formato = "EXCEL" then

            fxLeitorExcel(
                Content
            )

        else if Formato = "JSON" then

            fxLeitorJSON(
                Content
            )

        else if Formato = "XML" then

            fxLeitorXML(
                Content
            )

        else if Formato = "PDF" then

            fxLeitorPDF(
                Content
            )

        else

            error Error.Record(
                "Formato de arquivo não suportado",
                Text.Format(
                    "Não existe leitor implementado para o formato '#{0}'.",
                    {
                        Formato
                    }
                ),
                [
                    Formato = Formato
                ]
            )

in
    Resultado;

shared srcArquivos = let
    FonteDados =
        Text.Upper(
            fxParametro("Fonte_Dados")
        ),

    Resultado =
        if FonteDados <> "ARQUIVOS" then

            #table({}, {})

        else

            let
                Arquivos =
                    fxOrigemArquivos(),

                ArquivosFiltrados =
                    fxFiltrarArquivos(
                        Arquivos
                    ),

                Dados =
                    Table.AddColumn(
                        ArquivosFiltrados,
                        "Dados",
                        each
                            fxLeitorArquivo([Content]),
                        type any
                    )

            in
                Dados

in
    Resultado;

shared cfgParametrosFormatosArquivos = let
    Fonte =
        stgParametrosFormatosArquivos,

    Agrupado =
        Table.Group(
            Fonte,
            {"Name"},
            {
                {
                    "Value",
                    each List.Buffer([Value]),
                    type list
                }
            }
        ),

    Resultado =
        Record.FromTable(
            Agrupado
        )
in
    Resultado;

shared cfgParametrosSeveridades = let

    Fonte =
        Table.Buffer(
            stgParametrosSeveridades
        ),

    Registros =
        List.Combine(
            List.Transform(
                Table.ToRecords(Fonte),
                each
                    {
                        [
                            Nome = _[Descrição],
                            Valor =
                                [
                                    Código = _[Código],
                                    Descrição = _[Descrição],
                                    Ordem = _[Ordem],
                                    Bloqueia = _[Bloqueia]
                                ]
                        ],
                        [
                            Nome = _[Código],
                            Valor =
                                [
                                    Código = _[Código],
                                    Descrição = _[Descrição],
                                    Ordem = _[Ordem],
                                    Bloqueia = _[Bloqueia]
                                ]
                        ]
                    }
            )
        ),

    Resultado =
        Record.FromList(
            List.Transform(
                Registros,
                each [Valor]
            ),
            List.Transform(
                Registros,
                each [Nome]
            )
        )

in

    Resultado;

shared cfgCalendario = let
    DataInicialManual =
        fxParametro(
            "Data_Inicial_Calendario",
            null
        ),

    DataFinalManual =
        fxParametro(
            "Data_Final_Calendario",
            null
        ),

    IntervaloAutomatico =
        if DataInicialManual = null
            or DataFinalManual = null then

            fxCalendarioIntervalo(
                cfgIntervalosFatos,
                1,
                1
            )

        else
            null,

    DataInicial =
        if DataInicialManual <> null then
            Date.From(DataInicialManual)
        else
            IntervaloAutomatico[DataInicial],

    DataFinal =
        if DataFinalManual <> null then
            Date.From(DataFinalManual)
        else
            IntervaloAutomatico[DataFinal],

    _ =
        if DataInicial > DataFinal then
            error Error.Record(
                "Calendário",
                "A data inicial do calendário não pode ser posterior à data final.",
                [
                    DataInicial = DataInicial,
                    DataFinal = DataFinal
                ]
            )
        else
            null
in
    [
        DataInicial = DataInicial,
        DataFinal = DataFinal
    ];

shared cfgIntervalosFatos = // Estrutura utilizada para descobrir o intervalo de datas
// que será utilizado na dimensão de calendario.

{
    // Adicione uma entrada por tabela fato e sua coluna de data.
    fxIntervaloData(fatoVendas, "Data")
};

shared cfgObjetos = /*
------------------------------------------------------------------------------
Consulta.....: cfgObjetos

Descrição....:
Catálogo central dos objetos disponíveis no ambiente.

Agrupa os objetos normalizados por origem e disponibiliza acesso rápido
através de Records indexados pelo nome do objeto.

Estrutura:

    PowerQuery
        Record contendo os objetos do Power Query.

    Excel
        Record contendo os objetos do Excel.

------------------------------------------------------------------------------
*/

let

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

    Fonte =

        Table.Buffer(stgObjetos),

//--------------------------------------------------------------------------
// Função auxiliar
//--------------------------------------------------------------------------

    CriarCatalogo =

        (Origem as text) as record =>

            let

                Objetos =

                    Table.SelectRows(

                        Fonte,

                        each [Origem] = Origem

                    ),

                Registros =

                    Table.ToRecords(
                        Objetos
                    ),

                Resultado =

                    Record.FromList(

                        Registros,

                        Objetos[Nome]

                    )

            in

                Resultado,

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        [

            PowerQuery =

                CriarCatalogo(
                    "PowerQuery"
                ),

            Excel =

                CriarCatalogo(
                    "Excel"
                )

        ]
in
    Resultado

;

shared cfgFuncoesPowerQuery = let

//--------------------------------------------------------------------------
// Objetos
//--------------------------------------------------------------------------

    Objetos =

        List.Select(

            Record.FieldValues(

                cfgObjetos[PowerQuery]

            ),

            each

                [Kind] = "Function"

        ),

//--------------------------------------------------------------------------
// Funções
//--------------------------------------------------------------------------

    Funcoes =

        Record.FromList(

            List.Transform(

                Objetos,

                each

                    Record.Field(

                        #sections[Section1],

                        [Nome]

                    )

            ),

            List.Transform(

                Objetos,

                each [Nome]

            )

        )

in

    Funcoes;

shared fxFiltrarArquivos = (Tabela as table) as table =>

let
    Formato =
        Text.Upper(
            fxParametro("Formato_Arquivo")
        ),

    Extensoes =
        try
            Record.Field(
                cfgParametrosFormatosArquivos,
                Formato
            )
        otherwise
            error Error.Record(
                "Formato de arquivo inválido",
                Text.Format(
                    "O formato '#{0}' não está cadastrado em cfgFormatosArquivos.",
                    {
                        Formato
                    }
                ),
                [
                    Formato = Formato
                ]
            ),

    Resultado =
        Table.SelectRows(
            Tabela,
            each
                List.Contains(
                    Extensoes,
                    Text.Lower([Extension])
                )
        )

in
    Resultado;

shared stgDados = let
    Fonte = fxOrigem(),

    Tabela =
        fxComoTabela(
            Fonte
        ),

    // Remove linhas completamente vazias.
    LinhasValidas =
        Table.SelectRows(
            Tabela,
            each
                List.NonNullCount(
                    Record.FieldValues(_)
                ) > 0
        ),

    // Remove colunas que não serão utilizadas.
    ColunasRemovidas =
        LinhasValidas,

    // Renomeia colunas para o padrão do projeto.
    ColunasRenomeadas =
        ColunasRemovidas,

    // Garante a existência de colunas opcionais.
    ColunasGarantidas =
        ColunasRenomeadas,

    // Define os tipos de dados.
    Tipos =
        ColunasGarantidas,

    // Limpa espaços excedentes em colunas de texto.
    TextosPadronizados =
        Tipos,

    // Remove registros duplicados quando aplicável.
    RegistrosUnicos =
        TextosPadronizados,

    // Reordena as colunas para facilitar a leitura.
    ColunasReordenadas =
        RegistrosUnicos,

    Resultado =
        ColunasReordenadas

in
    Resultado;

shared srcREST = let
    FonteDados =
        Text.Upper(
            fxParametro("Fonte_Dados")
        ),

    Fonte =
        if FonteDados = "REST" then

            fxOrigemREST()

        else

            #table(
                {},
                {}
            )
in
    Fonte;

shared srcSGBD = let
    FonteDados =
        Text.Upper(
            fxParametro("Fonte_Dados")
        ),

    Fonte =
        if FonteDados = "SGBD" then

            fxOrigemSGBD()

        else

            #table(
                {},
                {}
            )

in
    Fonte;

shared parTabelaParametrosFormatosArquivos = "tbParametrosFormatosArquivos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaParametrosTratamentos = "tbParametrosTratamentos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaParametrosValidacoes = "tbParametrosValidacoes" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaParametrosSeveridades = "tbParametrosSeveridades" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaSchema = "tbSchema" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaClientes = "tbClientes" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaProdutos = "tbProdutos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaVendas = "tbVendas" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaDadosGenericos = "tbDadosGenericos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared nrmProdutos = let

    // Obtém os dados preparados na camada Stage.
    Fonte =
        stgProdutos,

    Normalizacao =
        fxNrmAplicar(
            Fonte,
            "tbProdutos"
        ),

    // Mantém apenas um registro para cada código de produto.
    RegistrosUnicos =
        Table.Distinct(
            Normalizacao,
            {"Código"}
        )

in

    RegistrosUnicos;

shared nrmVendas = let

    Fonte =
        stgVendas,

    Normalizacao =
        fxNrmAplicar(
            Fonte,
            "tbVendas"
        ),

    Clientes =
        List.Buffer(
            nrmClientes[CPF]
        ),

    Produtos =
        List.Buffer(
            nrmProdutos[Código]
        ),

    RegistrosRelacionados =
        Table.SelectRows(
            Normalizacao,
            each
                List.Contains(
                    Clientes,
                    [CPF]
                )
                and
                List.Contains(
                    Produtos,
                    [CódigoProduto]
                )
        ),

    Atributos =
        Table.AddColumn(
            RegistrosRelacionados,
            "ValorTotal",
            each
                [Quantidade] * [ValorUnitário],
            type number
        )

in

    Atributos;

shared nrmDados = let
    Fonte = stgDados
in
    Fonte;

shared dimProdutos = let
    // Obtém a entidade normalizada.
    Fonte =
        nrmProdutos,

    // Seleciona apenas os atributos da dimensão.
    Atributos =
        Table.SelectColumns(
            Fonte,
            {
                "Código",
                "Descrição",
                "Categoria",
                "PreçoLista"
            }
        ),

    // Remove registros duplicados.
    RegistrosUnicos =
        Table.Distinct(
            Atributos,
            {"Código"}
        ),

    // Ordena os registros para garantir estabilidade da chave substituta.
    RegistrosOrdenados =
        Table.Sort(
            RegistrosUnicos,
            {
                {"Código", Order.Ascending}
            }
        ),

    // Gera a chave substituta da dimensão.
    Chaves =
        Table.AddIndexColumn(
            RegistrosOrdenados,
            "IDProduto",
            1,
            1,
            Int64.Type
        ),

    // Reorganiza as colunas da dimensão.
    ColunasReordenadas =
        Table.ReorderColumns(
            Chaves,
            {
                "IDProduto",
                "Código",
                "Descrição",
                "Categoria",
                "PreçoLista"
            }
        ),
    
    Buffer = Table.Buffer(ColunasReordenadas)

in
    Buffer;

shared fxCalendarioBase = (
    DataInicial as date,
    DataFinal as date,
    optional PrimeiroDiaSemana as nullable number
) as table =>

let
    PrimeiroDia =
        if PrimeiroDiaSemana = null then
            Day.Monday
        else
            PrimeiroDiaSemana,

    _ =
        if DataInicial > DataFinal then
            error Error.Record(
                "Calendário",
                "A data inicial não pode ser posterior à data final.",
                [
                    DataInicial = DataInicial,
                    DataFinal = DataFinal
                ]
            )
        else
            null,

    QuantidadeDias =
        Duration.Days(DataFinal - DataInicial) + 1,

    Datas =
        List.Dates(
            DataInicial,
            QuantidadeDias,
            #duration(1, 0, 0, 0)
        ),

    Calendario =
        Table.FromColumns(
            {
                Datas,
                List.Transform(Datas, each Date.Year(_)),
                List.Transform(Datas, each if Date.Month(_) <= 6 then 1 else 2),
                List.Transform(Datas, each Date.QuarterOfYear(_)),
                List.Transform(Datas, each Date.Month(_)),
                List.Transform(Datas, each Date.Day(_)),
                List.Transform(Datas, each Date.DayOfYear(_)),
                List.Transform(Datas, each Date.DayOfWeek(_, PrimeiroDia)),
                List.Transform(Datas, each Date.WeekOfYear(_, PrimeiroDia)),
                List.Transform(Datas, each Date.WeekOfMonth(_, PrimeiroDia))
            },
            type table [
                Data = date,
                Ano = Int64.Type,
                Semestre = Int64.Type,
                Trimestre = Int64.Type,
                Mês = Int64.Type,
                Dia = Int64.Type,
                DiaAno = Int64.Type,
                DiaSemana = Int64.Type,
                SemanaAno = Int64.Type,
                SemanaMês = Int64.Type
            ]
        )
in
    Calendario;

shared fxCalendario = (
    Calendario as table,
    optional Cultura as nullable text
) as table =>

let
    Idioma =
        if Cultura = null then "pt-BR" else Cultura,

    AdicionarAtributos =
        Table.AddColumn(
            Calendario,
            "_Calendario",
            each
                let
                    Data = [Data],
                    Ano = [Ano],
                    Trimestre = [Trimestre],
                    Mes = [Mês],
                    DiaSemana = [DiaSemana],

                    NomeMes = Date.MonthName(Data, Idioma),
                    NomeDia = Date.DayOfWeekName(Data, Idioma)
                in
                    [
                        NomeMês = NomeMes,
                        NomeMêsAbrev = Text.Start(NomeMes, 3),
                        NomeDiaSemana = NomeDia,
                        NomeDiaSemanaAbrev = Text.Start(NomeDia, 3),
                        AnoMês =
                            Text.From(Ano)
                            & "-"
                            & Text.PadStart(Text.From(Mes), 2, "0"),
                        AnoTrimestre =
                            Text.From(Ano)
                            & " T"
                            & Text.From(Trimestre),
                        ÉDiaÚtil = DiaSemana < 5,
                        ÉFimSemana = DiaSemana >= 5,
                        PrimeiroDiaMês = Date.StartOfMonth(Data),
                        ÚltimoDiaMês = Date.EndOfMonth(Data),
                        PrimeiroDiaAno = Date.StartOfYear(Data),
                        ÚltimoDiaAno = Date.EndOfYear(Data)
                    ],
            type record
        ),

    Resultado =
        Table.ExpandRecordColumn(
            AdicionarAtributos,
            "_Calendario",
            {
                "NomeMês",
                "NomeMêsAbrev",
                "NomeDiaSemana",
                "NomeDiaSemanaAbrev",
                "AnoMês",
                "AnoTrimestre",
                "ÉDiaÚtil",
                "ÉFimSemana",
                "PrimeiroDiaMês",
                "ÚltimoDiaMês",
                "PrimeiroDiaAno",
                "ÚltimoDiaAno"
            }
        )
in
    Resultado;

shared fxCalendarioIntervalo = (
    intervalos as list,
    optional margemAnterior as nullable number,
    optional margemPosterior as nullable number
) as record =>

let
    AnosAntes =
        if margemAnterior = null then 1 else margemAnterior,

    AnosDepois =
        if margemPosterior = null then 1 else margemPosterior,

    IntervalosValidos =
        List.Select(
            List.RemoveNulls(intervalos),
            each
                [DataInicial] <> null
                and [DataFinal] <> null
        ),

    _ =
        if List.IsEmpty(IntervalosValidos) then
            error Error.Record(
                "Calendário",
                "Nenhum intervalo de datas válido foi informado."
            )
        else
            null,

    MenorData =
        List.Min(
            List.Transform(
                IntervalosValidos,
                each Date.From([DataInicial])
            )
        ),

    MaiorData =
        List.Max(
            List.Transform(
                IntervalosValidos,
                each Date.From([DataFinal])
            )
        )
in
    [
        DataInicial =
            Date.StartOfYear(
                Date.AddYears(MenorData, -AnosAntes)
            ),

        DataFinal =
            Date.EndOfYear(
                Date.AddYears(MaiorData, AnosDepois)
            )
    ];

shared dimCalendario = let

    Intervalo =
        cfgCalendario,

    CalendarioBase =

        fxCalendarioBase(

            Intervalo[DataInicial],

            Intervalo[DataFinal]

        ),

    Calendario =

        fxCalendario(

            CalendarioBase

        )

in

    Calendario;

shared fxSchema = (
    tabela as text
)
as record =>

let
    Resultado =
        Record.FieldOrDefault(
            cfgSchema,
            tabela,
            []
        )

in
    Resultado;

shared fxSchemaInterpretarTratamentos = (
    tratamento as text
)
as record =>

let
    Texto =
        Text.Trim(
            Text.From(tratamento)
        ),

    InicioParametros =
        Text.PositionOf(
            Texto,
            "("
        ),

    FimParametros =
        Text.PositionOf(
            Texto,
            ")"
        ),

    SintaxeValida =
        (InicioParametros = -1 and FimParametros = -1)
        or
        (
            InicioParametros >= 0
            and
            FimParametros > InicioParametros
        ),

    _ =
        if SintaxeValida then

            null

        else

            error
                "Sintaxe inválida no tratamento '" &
                Texto &
                "'. Utilize o formato NOME(parametro1,parametro2).",

    Nome =
        Text.Upper(

            Text.Trim(

                if InicioParametros = -1 then

                    Texto

                else

                    Text.Start(
                        Texto,
                        InicioParametros
                    )

            )

        ),

    ParametrosTexto =
        if InicioParametros = -1 then

            ""

        else

            Text.Trim(

                Text.BetweenDelimiters(
                    Texto,
                    "(",
                    ")"
                )

            ),

    Parametros =
        if ParametrosTexto = "" then

            {}

        else

            List.Transform(

                Text.Split(
                    ParametrosTexto,
                    ","
                ),

                each Text.Trim(_)

            )

in

    [
        Nome = Nome,
        Parametros = Parametros,
        Severidade = null,
        Ativo = true
    ];

shared fxSchemaTratamentoTrim = (
    valor as any,
    optional parametros as nullable any
)
as any =>

let

    Resultado =
        if valor = null then
            null
        else
            Text.Trim(
                Text.From(valor)
            )

in

    Resultado;

shared fxSchemaTratamentoUpper = (
    valor as any,
    optional parametros as nullable any
)
as any =>

let

    Resultado =
        if valor = null then
            null
        else
            Text.Upper(
                Text.From(valor)
            )

in

    Resultado;

shared fxSchemaTratamentoLower = (
    valor as any,
    optional parametros as nullable any
)
as any =>

let

    Resultado =
        if valor = null then
            null
        else
            Text.Lower(
                Text.From(valor)
            )

in

    Resultado;

shared fxSchemaTratamentoProper = (
    valor as any,
    optional parametros as nullable any
)
as any =>

let

    Resultado =
        if valor = null then
            null
        else
            Text.Proper(
                Text.From(valor)
            )

in

    Resultado
;

shared fxSchemaTratamentoClean = (
    valor as any,
    optional parametros as nullable any
)
as any =>

let

    Resultado =
        if valor = null then
            null
        else
            Text.Clean(
                Text.From(valor)
            )

in

    Resultado
;

shared fxSchemaTratamentoEmptyToNull = (
    valor as any,
    optional parametros as nullable any
)
as any =>

let
    Resultado =
        if valor = "" then
            null
        else
            valor
in
    Resultado;

shared fxSchemaTratamentoNullToEmpty = (
    valor as any,
    optional parametros as nullable any
)
as any =>

let
    Resultado =
        if valor = null then
            ""
        else
            valor
in
    Resultado;

shared fxSchemaTratamentoSingleSpace = (
    valor as any,
    optional parametros as nullable any
)
as any =>

let

    Resultado =
        if valor = null then
            null
        else
            Text.Combine(
                List.Select(
                    Text.SplitAny(
                        Text.Trim(
                            Text.From(valor)
                        ),
                        " "
                    ),
                    each _ <> ""
                ),
                " "
            )
in
    Resultado;

shared fxSchemaTratamentoDigits = (
    valor as any,
    optional parametros as nullable any
)
as any =>

let
    Resultado =
        if valor = null then
            null
        else
            Text.Select(
                Text.From(valor),
                {"0".."9"}
            )
in
    Resultado;

shared fxSchemaTratamentoAlphaNumeric = (
    valor as any,
    optional parametros as nullable any
)
as any =>

let
    Resultado =
        if valor = null then
            null
        else
            Text.Select(
                Text.From(valor),
                {"A".."Z","a".."z","0".."9"}
            )
in
    Resultado;

shared fxSchemaTratamentoAbs = (
    valor as any,
    optional parametros as nullable any
)
as any =>

let
    Resultado =
        if valor = null then
            null
        else
            Number.Abs(
                Number.From(valor)
            )
in
    Resultado;

shared fxSchemaTratamentoRound = (
    valor as any,
    optional parametros as nullable any
)
as any =>

let
    Resultado =
        if valor = null then
            null
        else
            Number.Round(
                Number.From(valor)
            )
in
    Resultado;

shared fxSchemaOcorrencia = (
    codigo as text,
    contexto as record,
    optional valorOriginal as any,
    optional valor as any,
    optional descricao as nullable text,
    optional detalhes as any
)
as record =>

let

    Operador =

        Record.FieldOrDefault(
            contexto,
            "Operador",
            []
        ),

    Resultado =

        [

            Codigo =
                codigo,

            Severidade =
                Record.FieldOrDefault(
                    Operador,
                    "Severidade"
                ),

            Coluna =
                Record.FieldOrDefault(
                    contexto,
                    "Coluna"
                ),

            Valor =
                valor,

            Tipo =

                Record.FieldOrDefault(
                    contexto,
                    "Tipo"
                ),

            Parametros =
                Record.FieldOrDefault(
                    Operador,
                    "Parametros"
                ),

            Mensagem =
                descricao,

            Detalhes =
                detalhes

        ]

in

    Resultado;

shared fxSchemaInterpretarValidacao = (
    validacao as text
)
as record =>

let
    Texto =
        Text.Trim(
            Text.From(validacao)
        ),

    InicioParametros =
        Text.PositionOf(
            Texto,
            "("
        ),

    FimParametros =
        Text.PositionOf(
            Texto,
            ")"
        ),

    SintaxeValida =
        (InicioParametros = -1 and FimParametros = -1)
        or
        (
            InicioParametros >= 0
            and
            FimParametros > InicioParametros
        ),

    _ =
        if SintaxeValida then

            null

        else

            error
                "Sintaxe inválida na validação '" &
                Texto &
                "'. Utilize o formato NOME(parametro1,parametro2).",

    Nome =
        Text.Upper(

            Text.Trim(

                if InicioParametros = -1 then

                    Texto

                else

                    Text.Start(
                        Texto,
                        InicioParametros
                    )

            )

        ),

    ParametrosTexto =
        if InicioParametros = -1 then

            ""

        else

            Text.Trim(

                Text.BetweenDelimiters(
                    Texto,
                    "(",
                    ")"
                )

            ),

    Parametros =
        if ParametrosTexto = "" then

            {}

        else

            List.Transform(

                Text.Split(
                    ParametrosTexto,
                    ","
                ),

                each Text.Trim(_)

            )

in

    [
        Nome = Nome,
        Parametros = Parametros,
        Severidade = null,
        Ativo = true
    ];

shared fxSchemaValidacaoList = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let

    Lista =
        if parametros = null then
            {}
        else
            parametros,

    Valido =

        if valor = null then

            true

        else

            List.Contains(
                Lista,
                Text.From(valor)
            ),

    Ocorrencias =

        if Valido then

            null

        else

            {

                fxSchemaOcorrencia(

                    "LIST",

                    contexto,

                    valor,

                    valor,

                    "Valor não pertence à lista permitida.",

                    [

                        Permitidos = Lista

                    ]

                )

            }

in

    [

        Valor = valor,

        Ocorrencias = Ocorrencias

    ];

shared fxSchemaValidacaoDomain = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let

    Dominio =
        if parametros = null then
            {}
        else
            List.Transform(
                parametros,
                each Text.Upper(Text.From(_))
            ),

    Valor =

        if valor = null then
            null
        else
            Text.Upper(
                Text.From(valor)
            ),

    Valido =

        if valor = null then
            true
        else
            List.Contains(
                Dominio,
                Valor
            ),

    Ocorrencias =

        if Valido then
            null
        else
            {
                fxSchemaOcorrencia(

                    "DOMAIN",

                    contexto,

                    valor,

                    valor,

                    "Valor fora do domínio permitido.",

                    [

                        Permitidos = parametros

                    ]

                )
            }

in

    [
        Valor = valor,
        Ocorrencias = Ocorrencias
    ];

shared fxSchemaValidacaoSize = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let

    TamanhoEsperado =

        if parametros = null or List.IsEmpty(parametros) then
            null
        else
            Number.From(parametros{0}),

    TamanhoAtual =

        if valor = null then
            null
        else
            Text.Length(
                Text.From(valor)
            ),

    Valido =

        valor = null
        or TamanhoEsperado = null
        or TamanhoAtual = TamanhoEsperado,

    Ocorrencias =

        if Valido then

            null

        else

            {

                fxSchemaOcorrencia(

                    "SIZE",

                    contexto,

                    valor,

                    valor,

                    "Tamanho inválido.",

                    [

                        Esperado = TamanhoEsperado,
                        Atual = TamanhoAtual

                    ]

                )

            }

in

    [

        Valor = valor,

        Ocorrencias = Ocorrencias

    ];

shared fxSchemaValidacaoMin = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let

    Minimo =

        if
            parametros = null
            or List.IsEmpty(parametros)
            or parametros{0} = null
        then
            null
        else
            try Number.From(parametros{0}) otherwise null,

    Valor =

        if valor = null then
            null
        else
            try Number.From(valor) otherwise null,

    Valido =

        if valor = null then

            true

        else

            Minimo <> null
            and Valor <> null
            and Valor >= Minimo,

    Ocorrencias =

        if Valido then

            null

        else

            {

                fxSchemaOcorrencia(

                    "MIN",

                    contexto,

                    valor,

                    valor,

                    "Valor inferior ao mínimo permitido.",

                    [

                        Minimo = Minimo,
                        Valor = Valor

                    ]

                )

            }

in

    [

        Valor = valor,

        Ocorrencias = Ocorrencias

    ];

shared fxSchemaValidacaoMax = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let

    Maximo =

        if
            parametros = null
            or List.IsEmpty(parametros)
            or parametros{0} = null
        then

            null

        else

            try
                Number.From(parametros{0})
            otherwise
                null,

    Valor =

        if valor = null then

            null

        else

            try
                Number.From(valor)
            otherwise
                null,

    Valido =

        if valor = null then

            true

        else if Maximo = null then

            true

        else if Valor = null then

            false

        else

            Valor <= Maximo,

    Ocorrencias =

        if Valido then

            null

        else

            {

                fxSchemaOcorrencia(

                    "MAX",

                    contexto,

                    valor,

                    valor,

                    "Valor superior ao máximo permitido.",

                    [

                        Maximo = Maximo,
                        Valor = Valor

                    ]

                )

            }

in

    [

        Valor = valor,

        Ocorrencias = Ocorrencias

    ];

shared fxSchemaValidacaoInterval = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let

    Minimo =

        if parametros = null
            or List.Count(parametros) < 1
            or parametros{0} = null then

            null

        else

            Number.From(parametros{0}),

    Maximo =

        if parametros = null
            or List.Count(parametros) < 2
            or parametros{1} = null then

            null

        else

            Number.From(parametros{1}),

    Valor =

        if valor = null then

            null

        else

            Number.From(valor),

    Valido =

        if Valor = null then

            true

        else

            (Minimo = null or Valor >= Minimo)
            and
            (Maximo = null or Valor <= Maximo),

    Ocorrencias =

        if Valido then

            null

        else

            {

                fxSchemaOcorrencia(

                    "INTERVAL",

                    contexto,

                    valor,

                    valor,

                    "Valor fora do intervalo permitido.",

                    [

                        Minimo = Minimo,
                        Maximo = Maximo,
                        Valor = Valor

                    ]

                )

            }

in

    [

        Valor = valor,
        Ocorrencias = Ocorrencias

    ];

shared fxSchemaValidacaoEmail = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let

    Valido =

        if valor = null then

            true

        else

            let

                Texto =
                    Text.From(valor),

                Partes =
                    Text.Split(
                        Texto,
                        "@"
                    )

            in

                List.Count(Partes) = 2
                and Text.Length(Partes{0}) > 0
                and Text.Contains(Partes{1}, ".")
                and Text.PositionOf(Partes{1}, ".") > 0
                and Text.PositionOf(Partes{1}, ".") < Text.Length(Partes{1}) - 1,

    Ocorrencias =

        if Valido then

            null

        else

            {

                fxSchemaOcorrencia(

                    "EMAIL",

                    contexto,

                    valor,

                    valor,

                    "E-mail inválido.",

                    []

                )

            }

in

    [

        Valor = valor,

        Ocorrencias = Ocorrencias

    ];

shared fxSchemaValidacaoURL = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let

    URL =

        if valor = null then

            null

        else

            Text.Trim(
                Text.From(valor)
            ),

    Protocolo =

        if URL = null then

            null

        else if Text.StartsWith(URL, "http://") then

            "http://"

        else if Text.StartsWith(URL, "https://") then

            "https://"

        else

            null,

    Dominio =

        if Protocolo = null then

            ""

        else

            Text.AfterDelimiter(
                URL,
                Protocolo
            ),

    Valido =

        URL = null

        or

        (

            Protocolo <> null

            and Text.Contains(Dominio, ".")

            and Text.PositionOf(Dominio, ".") > 0

            and Text.PositionOf(Dominio, ".") < Text.Length(Dominio) - 1

        ),

    Ocorrencias =

        if Valido then

            null

        else

            {

                fxSchemaOcorrencia(

                    "URL",

                    contexto,

                    valor,

                    valor,

                    "Endereço URL inválido.",

                    []

                )

            }

in

    [

        Valor = valor,

        Ocorrencias = Ocorrencias

    ];

shared fxSchemaValidacaoCEP = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let

    Ocorrencias =

        if valor = null then

            null

        else

            let

                CEP =
                    Text.Trim(
                        Text.From(valor)
                    ),

                Digitos =
                    Text.Select(
                        CEP,
                        {"0".."9"}
                    ),

                FormatoSemMascara =
                    Text.Length(CEP) = 8
                    and
                    Text.Length(Digitos) = 8,

                FormatoComMascara =
                    Text.Length(CEP) = 9
                    and
                    Text.Length(Digitos) = 8
                    and
                    Text.Middle(CEP, 5, 1) = "-",

                Valido =
                    FormatoSemMascara
                    or
                    FormatoComMascara

            in

                if Valido then

                    null

                else

                    {

                        fxSchemaOcorrencia(

                            "CEP",

                            contexto,

                            valor,

                            valor,

                            "CEP inválido.",

                            []

                        )

                    }

in

    [

        Valor = valor,

        Ocorrencias = Ocorrencias

    ];

shared fxSchemaValidacaoCPF = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let

    CPF =
        if valor = null then
            null
        else
            Text.From(valor),

    PodeValidar =
        CPF <> null
        and
        Text.Length(CPF) = 11,

    Digitos =
        if PodeValidar then
            List.Transform(
                Text.ToList(CPF),
                each Number.FromText(_)
            )
        else
            {},

    NaoRepetido =
        PodeValidar
        and
        List.Count(
            List.Distinct(Digitos)
        ) > 1,

    SomaDV1 =
        if NaoRepetido then

            List.Sum(

                List.Transform(
                    {0..8},
                    each Digitos{_} * (10 - _)
                )

            )

        else

            null,

    RestoDV1 =
        if NaoRepetido then
            Number.Mod(SomaDV1 * 10, 11)
        else
            null,

    DV1 =
        if RestoDV1 = 10 then
            0
        else
            RestoDV1,

    SomaDV2 =
        if NaoRepetido then

            List.Sum(

                List.Transform(

                    {0..9},

                    each
                        (
                            if _ = 9 then
                                DV1
                            else
                                Digitos{_}
                        )
                        *
                        (11 - _)

                )

            )

        else

            null,

    RestoDV2 =
        if NaoRepetido then
            Number.Mod(SomaDV2 * 10, 11)
        else
            null,

    DV2 =
        if RestoDV2 = 10 then
            0
        else
            RestoDV2,

    Valido =

        if valor = null then

            true

        else if not PodeValidar then

            false

        else if not NaoRepetido then

            false

        else

            DV1 = Digitos{9}
            and
            DV2 = Digitos{10},

    Ocorrencias =

        if Valido then

            null

        else

            {
                fxSchemaOcorrencia(

                    "CPF",

                    contexto,

                    valor,

                    valor,

                    "CPF inválido.",

                    []

                )
            }

in

    [
        Valor = valor,
        Ocorrencias = Ocorrencias
    ];

shared fxSchemaValidacaoCNPJ = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let

    ValorNulo =
        valor = null,

    CNPJ =
        if ValorNulo then
            null
        else
            Text.Select(
                Text.From(valor),
                {"0".."9"}
            ),

    Digitos =
        if ValorNulo then
            {}
        else
            List.Transform(
                Text.ToList(CNPJ),
                each Number.FromText(_)
            ),

    TamanhoValido =
        List.Count(Digitos) = 14,

    NaoRepetido =
        TamanhoValido
        and
        List.Count(
            List.Distinct(Digitos)
        ) > 1,

    PesosDV1 =
        {5,4,3,2,9,8,7,6,5,4,3,2},

    SomaDV1 =
        if NaoRepetido then

            List.Sum(

                List.Transform(

                    {0..11},

                    each
                        Digitos{_} *
                        PesosDV1{_}

                )

            )

        else

            null,

    RestoDV1 =
        if NaoRepetido then
            Number.Mod(
                SomaDV1,
                11
            )
        else
            null,

    DV1 =
        if NaoRepetido then

            if RestoDV1 < 2 then
                0
            else
                11 - RestoDV1

        else

            null,

    PesosDV2 =
        {6,5,4,3,2,9,8,7,6,5,4,3,2},

    SomaDV2 =
        if NaoRepetido then

            List.Sum(

                List.Transform(

                    {0..12},

                    each

                        (
                            if _ = 12 then
                                DV1
                            else
                                Digitos{_}
                        )

                        *

                        PesosDV2{_}

                )

            )

        else

            null,

    RestoDV2 =
        if NaoRepetido then
            Number.Mod(
                SomaDV2,
                11
            )
        else
            null,

    DV2 =
        if NaoRepetido then

            if RestoDV2 < 2 then
                0
            else
                11 - RestoDV2

        else

            null,

    Valido =

        if ValorNulo then

            true

        else

            NaoRepetido
            and
            DV1 = Digitos{12}
            and
            DV2 = Digitos{13},

    Ocorrencias =

        if Valido then

            null

        else

            {

                fxSchemaOcorrencia(

                    "CNPJ",

                    contexto,

                    valor,

                    valor,

                    "CNPJ inválido.",

                    []

                )

            }

in

    [

        Valor = valor,

        Ocorrencias = Ocorrencias

    ];

shared tststgPreparar = let

    Fonte = srcClientes,

    Resultado =
        fxStgPreparar(
            Fonte
        )
in
    Resultado;

shared tstGeral = let

Fonte = srcClientes,
    #"Linhas em Branco Removidas" = Table.SelectRows(Fonte, each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null})))

in
#"Linhas em Branco Removidas";

shared fxIntervaloData = (
    tabela as table,
    coluna as text
) as nullable record =>

let
    // MUDANÇA DE ABORDAGEM:
    // Anteriormente, o motor lia a coluna inteira, criava uma lista em memória e executava try/catch em todas
    // as linhas para converter em Data.
    // Para otimizar, primeiro inspecionamos o tipo estático da coluna na tabela com Type.TableColumn.
    // Se a coluna já estiver tipada no modelo como Date ou DateTime, nós simplesmente extraímos a coluna e removemos
    // os nulos, eliminando o loop de conversão de dados por completo.
    // Se não estiver tipada, filtramos os valores em branco/nulos primeiro (reduzindo o volume de processamento)
    // e fazemos a conversão apenas para os tipos que realmente necessitam.
    SchemaTipo = try Type.TableColumn(Value.Type(tabela), coluna) otherwise type any,
    ValoresOrigem = Table.Column(tabela, coluna),
    
    Datas = 
        if SchemaTipo = type date or SchemaTipo = type datetime or SchemaTipo = type datetimezone then
            List.RemoveNulls(ValoresOrigem)
        else
            let
                // Seleciona apenas os valores não vazios para evitar processar nulos/vazios
                NaoNulos = List.Select(ValoresOrigem, each _ <> null and _ <> ""),
                Convertidos = List.Transform(NaoNulos, (v) => 
                    if v is date then v 
                    else if v is datetime then DateTime.Date(v) 
                    else try Date.From(v) otherwise null
                )
            in
                List.RemoveNulls(Convertidos)
in
    if List.IsEmpty(Datas) then
        null
    else
        [
            DataInicial = List.Min(Datas),
            DataFinal = List.Max(Datas)
        ]
;

shared fxSchemaValidacaoREQUIRED = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let

    Vazio =

        if valor = null then

            true

        else if valor is text then

            Text.Trim(valor) = ""

        else

            false,

    Ocorrencias =

        if Vazio then

            {

                fxSchemaOcorrencia(

                    "REQUIRED",

                    contexto,

                    valor,

                    valor,

                    "Campo obrigatório não informado.",

                    []

                )

            }

        else

            null

in

    [

        Valor =
            valor,

        Ocorrencias =
            Ocorrencias

    ];

shared fxNrmDescartarRegistrosBloqueantes = (
    tabela as table
)
as table =>

let

    TemOcorrencias =
        Table.HasColumns(
            tabela,
            "Ocorrencias"
        ),

    TabelaFiltrada =

        if
            not TemOcorrencias
        then

            tabela

        else

            Table.SelectRows(

                tabela,

                each

                    let

                        Ocorrencias =

                            if [Ocorrencias] = null then
                                {}
                            else
                                [Ocorrencias]

                    in

                        not List.MatchesAny(

                            Ocorrencias,

                            (Ocorrencia) =>

                                Record.FieldOrDefault(

                                    Record.FieldOrDefault(

                                        cfgParametrosSeveridades,

                                        Ocorrencia[Severidade],

                                        []

                                    ),

                                    "Bloqueia",

                                    false

                                )

                        )

            ),

    Resultado =
        Table.RemoveColumns(
            TabelaFiltrada,
            "Ocorrencias",
            MissingField.Ignore
        )

in

    Resultado;

shared parTabelaCategoriasConsultasPQ = "tbSobreCategoriasConsultasPQ" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared fxStgPreparar = /*
Responsabilidades:
- Preparar a estrutura mínima de uma tabela antes da aplicação do Schema.
- Remover registros completamente vazios considerando apenas as colunas efetivas.
- Garantir que a tabela possua pelo menos uma coluna.
- Validar os nomes das colunas após a normalização:
    - não vazios;
    - únicos.
- Normalizar os nomes das colunas removendo espaços nas extremidades.
- Permitir excluir colunas específicas da preparação por meio de "ignorarColunas".
- Preservar integralmente os valores das células.
- Não adicionar, remover ou alterar colunas.
- Não depender do Schema nem de configurações externas.
*/

(
    tabela as table,
    optional ignorarColunas as nullable list
)
as table =>

let

    ColunasIgnoradas =
        if ignorarColunas = null then
            {}
        else
            ignorarColunas,

    // Obtém os nomes atuais das colunas.
    ColunasTabela =
        Table.ColumnNames(
            tabela
        ),

    // Garante que exista pelo menos uma coluna.
    _ErroSemColunas =
        if List.IsEmpty(ColunasTabela) then
            error "A tabela não possui colunas."
        else
            null,

    // Remove das colunas da tabela aquelas que não devem participar da preparação.
    ColunasEfetivas =
        List.Difference(
            ColunasTabela,
            ColunasIgnoradas
        ),

    // Remove linhas completamente vazias considerando apenas as colunas efetivas.
    TabelaSemLinhasVazias =
        if List.IsEmpty(ColunasEfetivas) then
            tabela
        else
            Table.SelectRows(
                tabela,
                each
                    not List.IsEmpty(
                        List.RemoveMatchingItems(
                            Record.FieldValues(
                                Record.SelectFields(
                                    _,
                                    ColunasEfetivas,
                                    MissingField.Ignore
                                )
                            ),
                            {"", null}
                        )
                    )
            ),

    // Normaliza apenas os nomes das colunas efetivas.
    ColunasNormalizadas =
        List.ReplaceMatchingItems(
            ColunasTabela,
            List.Transform(
                ColunasEfetivas,
                each {_, Text.Trim(_)}
            )
        ),

    // Verifica nomes vazios após a normalização.
    TemNomeInvalido =
        List.AnyTrue(
            List.Transform(
                ColunasNormalizadas,
                each _ = ""
            )
        ),

    _ErroNomeInvalido =
        if TemNomeInvalido then
            error "A tabela possui colunas com nome vazio."
        else
            null,

    // Verifica duplicidade após a normalização.
    TemDuplicidade =
        List.Count(ColunasNormalizadas)
            <>
        List.Count(
            List.Distinct(
                ColunasNormalizadas
            )
        ),

    _ErroDuplicidade =
        if TemDuplicidade then
            error "A tabela possui nomes de colunas duplicados."
        else
            null,

    // Aplica a normalização dos nomes das colunas.
    Resultado =
        Table.RenameColumns(
            TabelaSemLinhasVazias,
            List.Zip(
                {
                    ColunasTabela,
                    ColunasNormalizadas
                }
            ),
            MissingField.Ignore
        )

in

    Resultado;

shared fxStgGarantirColunas = /*
Responsabilidades:
- Garantir que todas as colunas do Pipeline existam.
- Adicionar apenas as colunas faltantes.
- Nunca remover colunas.
- Nunca alterar tipos.
- Nunca ordenar colunas.
*/

(
    Tabela as table,
    Pipeline as record
)
as table =>

let

//--------------------------------------------------------------------------
// Colunas
//--------------------------------------------------------------------------

    ColunasPipeline =
        Pipeline[Colunas],

    ColunasExistentes =
        Table.ColumnNames(
            Tabela
        ),

    ColunasFaltantes =
        List.Difference(
            ColunasPipeline,
            ColunasExistentes
        ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        List.Accumulate(

            ColunasFaltantes,

            Tabela,

            (
                Estado,
                Coluna
            ) =>

                Table.AddColumn(

                    Estado,

                    Coluna,

                    each null

                )

        )

in

    Resultado;

shared fxNrmPadronizar = /*
Responsabilidades:
- Aplicar os tratamentos definidos no Pipeline.
- Nunca aplicar validações.
- Nunca converter tipos.
- Nunca adicionar ou remover colunas.
*/

(
    tabela as table,
    pipeline as record
)
as table =>

let

//--------------------------------------------------------------------------
// Tratamentos
//--------------------------------------------------------------------------

    Tratamentos =
        pipeline[Tratamentos],

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        if List.IsEmpty(Tratamentos) then

            tabela

        else

            let

//--------------------------------------------------------------------------
// Estado inicial
//--------------------------------------------------------------------------

                NomesColunas =
                    Table.ColumnNames(tabela),

                IndicesColunas =
                    Record.FromList(
                        List.Positions(NomesColunas),
                        NomesColunas
                    ),

                EstadoInicial =

                    [

                        Colunas =
                            Table.ToColumns(tabela)

                    ],

//--------------------------------------------------------------------------
// Aplicação dos tratamentos
//--------------------------------------------------------------------------

                EstadoFinal =

                    List.Accumulate(

                        Tratamentos,

                        EstadoInicial,

                        (
                            estado,
                            tratamento
                        ) =>

                            let

                                Posicao =
                                    Record.Field(
                                        IndicesColunas,
                                        tratamento[Coluna]
                                    ),

                                NovosValores =

                                    fxStgAplicarTratamentos(

                                        estado[Colunas]{Posicao},

                                        tratamento[Operadores]

                                    ),

                                NovasColunas =

                                    List.ReplaceRange(

                                        estado[Colunas],

                                        Posicao,

                                        1,

                                        {

                                            NovosValores

                                        }

                                    )

                            in

                                [

                                    Colunas =
                                        NovasColunas

                                ]

                    ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

                ResultadoFinal =

                    Table.FromColumns(

                        EstadoFinal[Colunas],

                        NomesColunas

                    )

            in

                ResultadoFinal

in

    Resultado;

shared cfgTabelasExcel = /*
------------------------------------------------------------------------------
Consulta.....: cfgTabelasExcel

Descrição....:
Catálogo indexado das tabelas do Excel.

Cada campo do record representa uma tabela do Workbook, permitindo acesso
direto aos seus metadados pelo nome.

------------------------------------------------------------------------------
*/

let

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

    Fonte =

        Table.Buffer(stgTabelasExcel),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        Record.FromList(

            Table.ToRecords(
                Fonte
            ),

            Fonte[Nome]

        )

in

    Resultado;

shared cfgTabelasPowerQuery = /*
------------------------------------------------------------------------------
Consulta.....: cfgTabelasPowerQuery

Descrição....:
Catálogo indexado das tabelas do Power Query.

Cada campo do record representa uma consulta do tipo tabela, permitindo
acesso direto aos seus metadados pelo nome.

------------------------------------------------------------------------------
*/

let

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

    Fonte =

        Table.Buffer(stgTabelasPowerQuery),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        Record.FromList(

            Table.ToRecords(
                Fonte
            ),

            Fonte[Nome]

        )

in

    Resultado;

shared cfgTabelas = let

    Tabelas =

        [

            PowerQuery =

                cfgTabelasPowerQuery,

            Excel =

                cfgTabelasExcel

        ]
in
    Tabelas;

shared fxSchemaTratamentoNormalizeBasic = /*
------------------------------------------------------------------------------
Função......: fxSchemaTratamentoNormalizeBasic

Descrição...:
Normaliza valores textuais aplicando, em uma única operação lógica, os
tratamentos mais comuns de higienização.

Tratamentos executados:

    • TRIM
        Remove espaços no início e no final.

    • CLEAN
        Remove caracteres de controle não imprimíveis.

    • SINGLESPACE
        Substitui múltiplos espaços internos por um único espaço.

    • EMPTYTONULL
        Converte texto vazio em null.

Características:

    • Idempotente.
    • Não altera valores nulos.
    • Realiza apenas uma conversão para texto.
    • Não depende de contexto nem de parâmetros.

------------------------------------------------------------------------------
*/

(
    valor as any,
    optional parametros as nullable any
)
as any =>

let

    Resultado =

        if valor = null then

            null

        else

            let

                Texto =

                    Text.From(
                        valor
                    ),

                Limpo =

                    Text.Clean(
                        Texto
                    ),

                Aparado =

                    Text.Trim(
                        Limpo
                    ),

                Normalizado =

                    Text.Combine(

                        List.Select(

                            Text.SplitAny(
                                Aparado,
                                " "
                            ),

                            each _ <> ""

                        ),

                        " "

                    )

            in

                if Normalizado = "" then
                    null
                else
                    Normalizado

in

    Resultado;

shared fxIdentificarTipoObjeto = /*
------------------------------------------------------------------------------
Função.......: fxIdentificarTipoObjeto

Descrição....:
Identifica as características de um tipo M.

Parâmetros...:

    tipo
        Tipo retornado por Value.Type().

Retorno......:

    Record contendo todas as propriedades definidas em cfgTiposObjetos,
    acrescidas dos indicadores Is<Kind> para todos os tipos conhecidos.

Observações..:

    A função é completamente orientada pela consulta cfgTiposObjetos.
    Qualquer propriedade adicionada ao catálogo passa a ser suportada
    automaticamente, sem necessidade de alterações nesta função.

------------------------------------------------------------------------------
*/

(
    tipo as type
)
as record =>

let

    Tipos =
        Record.FieldValues(
            cfgTiposObjetos
        ),

    TiposComMatch =

        List.Transform(

            Tipos,

            each

                Record.AddField(

                    _,

                    "IsMatch",

                    Type.Is(
                        tipo,
                        _[Type]
                    )

                )

        ),

    InformacoesTipo =

        let

            Correspondencias =

                List.Select(

                    TiposComMatch,

                    each [IsMatch]

                )

        in

            if List.IsEmpty(Correspondencias)
            then
                [
                    Kind = "Unknown",
                    Type = tipo
                ]
            else
                Record.RemoveFields(

                    Correspondencias{0},

                    {"IsMatch"}

                ),

    Flags =

        Record.FromList(

            List.Transform(

                TiposComMatch,

                each [IsMatch]

            ),

            List.Transform(

                TiposComMatch,

                each "Is" & [Kind]

            )

        ),

    Retorno =

        Record.Combine(

            {
                InformacoesTipo,
                Flags
            }

        )

in

    Retorno;

shared fxIdentificarObjetoPeloValor = /*
------------------------------------------------------------------------------
Função.......: fxIdentificarObjeto

Descrição....:
Identifica as principais características de um objeto do Power Query.

Parâmetros...:

    valor
        Objeto a ser analisado.

Retorno......:

    Record contendo:

        Type
        Kind
        Category
        IsStructured
        ...
        IsTable
        IsRecord
        IsFunction
        IsList
        ...

Observações..:

    A classificação do tipo é delegada para a função
    fxIdentificarTipoObjeto().

    A identificação de parâmetros é realizada pela função
    fxParametrosIdentificarParametroPQ().

------------------------------------------------------------------------------
*/

(
    valor as any
)
as record =>

let

    Tipo =
        Value.Type(
            valor
        ),

    TipoInfo =
        fxIdentificarTipoObjeto(
            Tipo
        ),

    Resultado =
        Record.Combine(
            {
                [
                    Type = Tipo
                ],
                TipoInfo
            }
        )

in

    Resultado;

shared fxParametrosIdentificarParametroPQ = /*
------------------------------------------------------------------------------
Função.......: fxParametrosIdentificarParametroPQ

Descrição....:
Identifica se um objeto do Power Query corresponde a uma consulta definida
como parâmetro.

Parâmetros...:

    valor
        Objeto do Power Query.

Retorno......:

    true  -> Consulta é um parâmetro.
    false -> Caso contrário.

------------------------------------------------------------------------------
*/

(
    valor as any
)
as logical =>

let

    Metadata =

        try
            Value.Metadata(
                valor
            )
        otherwise
            [],

    Resultado =

        Record.FieldOrDefault(

            Metadata,

            "IsParameterQuery",

            false

        )

in

    Resultado;

shared fxStgAplicarTratamentos = (
    Valores as list,
    Tratamentos as list
)
as list =>

let

    Resultado =

        List.Accumulate(

            Tratamentos,

            Valores,

            (Estado, Operador) =>

                List.Transform(

                    Estado,

                    each

                        Operador[Função](

                            _,

                            Operador[Parametros]

                        )

                )

        )

in

    Resultado;

shared fxStgAplicarTipos = /*
Responsabilidades:
- Aplicar os tipos definidos no Pipeline.
- Inferir automaticamente o tipo quando o Pipeline informar type any.
- Nunca aplicar tratamentos.
- Nunca aplicar validações.
- Nunca adicionar, remover ou ordenar colunas.
*/

(
    Tabela as table,
    Pipeline as record,
    optional Amostra as nullable number
)
as table =>

let

//--------------------------------------------------------------------------
// Tipos
//--------------------------------------------------------------------------

    Tipos =
        Pipeline[Tipos],

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        if List.IsEmpty(Tipos) then

            Tabela

        else

            let

                Transformacoes =

                    List.Transform(

                        Tipos,

                        (Definicao) =>

                            let

                                Coluna =
                                    Definicao[Coluna],

                                TipoConfigurado =
                                    Definicao[Tipo],

                                TipoFinal =

                                    if TipoConfigurado = type any then

                                        fxStgIdentificarTipoColuna(

                                            Tabela,
                                            Coluna,
                                            Amostra

                                        )

                                    else

                                        TipoConfigurado

                            in

                                {

                                    Coluna,

                                    (Valor) =>

                                        if Valor = null then

                                            null

                                        else if Value.Is(

                                            Valor,
                                            TipoFinal

                                        ) then

                                            Valor

                                        else

                                            try

                                                fxConversor(

                                                    Valor,
                                                    TipoFinal

                                                )

                                            otherwise

                                                null,

                                    TipoFinal

                                }

                    )

            in

                Table.TransformColumns(

                    Tabela,

                    Transformacoes,

                    null,

                    MissingField.Ignore

                )

in

    Resultado;

shared fxStgIdentificarTipoColuna = /*
------------------------------------------------------------------------------
Função......: fxSchemaTipoColuna

Descrição...:
Retorna o tipo efetivo de uma coluna.

Funcionamento:

    1. Obtém o tipo declarado da coluna.

    2. Caso o tipo declarado seja diferente de 'type any',
       retorna imediatamente esse tipo.

    3. Caso o tipo seja 'type any', analisa uma pequena amostra
       de valores não nulos.

    4. Se todos os valores da amostra possuem o mesmo tipo,
       retorna esse tipo.

    5. Caso existam tipos diferentes ou a coluna esteja vazia,
       retorna 'type any'.

Parâmetros..:

    tabela
        Tabela contendo a coluna.

    coluna
        Nome da coluna.

    amostra (opcional)
        Quantidade máxima de valores não nulos utilizados para
        inferência do tipo.

        Padrão: 20

Retorno.....:

    Um valor do tipo 'type'.

Exemplos....:

    type text
    type number
    type date
    type datetime
    type logical
    type binary
    type any

------------------------------------------------------------------------------
*/

(
    tabela as table,
    coluna as text,
    optional amostra as nullable number
)
as type =>

let

    QuantidadeAmostra =
        if amostra = null then
            20
        else
            Number.RoundDown(amostra),

    TipoDeclarado =
        Type.TableColumn(
            Value.Type(tabela),
            coluna
        ),

    Resultado =

        if TipoDeclarado <> type any then

            TipoDeclarado

        else

            let

                Valores =

                    List.FirstN(

                        List.RemoveNulls(

                            Table.Column(
                                tabela,
                                coluna
                            )

                        ),

                        QuantidadeAmostra

                    ),

                Tipos =

                    List.Distinct(

                        List.Transform(

                            Valores,

                            each Value.Type(_)

                        )

                    )

            in

                if List.Count(Tipos) = 1 then
                    Tipos{0}
                else
                    type any

in

    Resultado;

shared fxIdentificarObjetoPeloNome = /*
------------------------------------------------------------------------------
Função......: fxIdentificarObjetoPeloNome

Descrição...:
Identifica um objeto do ambiente utilizando exclusivamente seu nome e as
configurações cadastradas no framework.

Não materializa objetos, evitando referências cíclicas.

Parâmetros..:

    objeto
        Nome do objeto.

    source
        Origem do objeto ("PowerQuery" ou "Excel").

Retorno.....:

    Record contendo os metadados do objeto.

------------------------------------------------------------------------------
*/

(
    objeto as text,
    source as text
)
as record =>

let

//--------------------------------------------------------------------------
// Prefixos cadastrados
//--------------------------------------------------------------------------

    Prefixos =

        Record.FieldNames(

            cfgCategoriasPowerQuery

        ),

//--------------------------------------------------------------------------
// Prefixo identificado
//--------------------------------------------------------------------------

    Prefixo =

        if source = "Excel" then

            ""

        else

            List.First(

                List.Select(

                    Prefixos,

                    each

                        Text.StartsWith(

                            Text.Lower(objeto),

                            Text.Lower(_)

                        )

                ),

                null

            ),

//--------------------------------------------------------------------------
// Categoria
//--------------------------------------------------------------------------

    Categoria =

        if source = "Excel" then

            [

                Saída = "Tabela"

            ]

        else

            Record.FieldOrDefault(

                cfgCategoriasPowerQuery,

                Prefixo,

                [

                    Saída = "Qualquer"

                ]

            ),

//--------------------------------------------------------------------------
// Definição do tipo
//--------------------------------------------------------------------------

    DefinicaoTipo =

        List.First(

            List.Select(

                Record.FieldValues(

                    cfgTiposObjetos

                ),

                each [Nome] = Categoria[Saída]

            ),

            null

        ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        if DefinicaoTipo = null then

            error Error.Record(

                "Tipo inválido",

                Text.Format(

                    "A saída '#{0}' não está cadastrada em cfgTiposObjetos.",

                    {

                        Categoria[Saída]

                    }

                )

            )

        else

            [

                Prefix = Prefixo,

                Kind = DefinicaoTipo[Kind],

                Tipo = DefinicaoTipo[Nome],

                Type = DefinicaoTipo[Type],

                Categoria = DefinicaoTipo[Categoria],

                IsStructured = DefinicaoTipo[IsStructured]

            ]

in

    Resultado;

shared tstPadronizar = let

    Fonte = srcClientes,

    Pipeline =
            fxPipeline("tbClientes"),

    Resultado =
        fxNrmPadronizar(
            Fonte,
            Pipeline
        )
in
    Resultado;

shared tstAplicarTipos = let

    Fonte =
        srcClientes,

    Pipeline =
        fxPipeline("tbClientes"),

    Resultado =
        fxStgAplicarTipos(
            Fonte,
            Pipeline
        )

in

    Resultado;

shared fxStgAplicar = (
    Tabela as table,
    optional Schema as nullable text
)
as table =>

let

//--------------------------------------------------------------------------
// Pipeline
//--------------------------------------------------------------------------

    Pipeline =

        if Schema = null then

            null

        else

            Record.Field(
                cfgPipeline,
                Schema
            ),

    TemPipeline =
        Pipeline <> null,

//--------------------------------------------------------------------------
// Preparação
//--------------------------------------------------------------------------

    Preparada =

        fxStgPreparar(
            Tabela
        ),

//--------------------------------------------------------------------------
// Pipeline
//--------------------------------------------------------------------------

    ColunasGarantidas =

        if TemPipeline then

            fxStgGarantirColunas(
                Preparada,
                Pipeline
            )

        else

            Preparada,

    ColunasRemovidas =

        if TemPipeline then

            fxStgRemoverColunas(
                ColunasGarantidas,
                Pipeline
            )

        else

            ColunasGarantidas,

    Tipada =

        if TemPipeline then

            fxStgAplicarTipos(
                ColunasRemovidas,
                Pipeline
            )

        else

            ColunasRemovidas,

    Resultado =

        if TemPipeline then

            fxStgOrdenarColunas(
                Tipada,
                Pipeline
            )

        else

            Tipada

in

    Resultado;

shared fxNrmAplicar = (
    Tabela as table,
    Schema as text
)
as table =>

let

//--------------------------------------------------------------------------
// Pipeline
//--------------------------------------------------------------------------

    Pipeline =

        Record.Field(
            cfgPipeline,
            Schema
        ),

//--------------------------------------------------------------------------
// Configuração
//--------------------------------------------------------------------------

    ProcessarTratamentos =

        fxParametro(
            "Processar_Tratamentos"
        ),

    ProcessarValidacoes =

        fxParametro(
            "Processar_Validacoes"
        ),

//--------------------------------------------------------------------------
// Pipeline
//--------------------------------------------------------------------------

    Tratada =

        if ProcessarTratamentos then

            fxNrmPadronizar(
                Tabela,
                Pipeline
            )

        else

            Tabela,

    Validada =

        if ProcessarValidacoes then

            fxNrmAplicarValidacoes(
                Tratada,
                Pipeline
            )

        else

            Tratada,

    Resultado =

        fxNrmDescartarRegistrosBloqueantes(
            Validada
        )

in

    Resultado;

shared tstStgAplicar = let

    Fonte = srcClientes,

    Schema = "tbClientes",

    Resultado =
        fxStgAplicar(
            Fonte,
            Schema
        )
in
    Resultado;

shared fxNrmAplicarValidacoes = /*
Responsabilidades:
- Executar as validações definidas no Pipeline.
- Nunca alterar os valores das colunas.
- Nunca converter tipos.
- Nunca adicionar ou remover linhas.
- Adicionar apenas a coluna "Ocorrencias".
*/

(
    Tabela as table,
    Pipeline as record
)
as table =>

let

//--------------------------------------------------------------------------
// Validações
//--------------------------------------------------------------------------

    Validacoes =
        Pipeline[Validações],

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        if List.IsEmpty(Validacoes) then

            Table.AddColumn(
                Tabela,
                "Ocorrencias",
                each null,
                type nullable list
            )

        else

            let

//--------------------------------------------------------------------------
// Estado inicial
//--------------------------------------------------------------------------

                NomesColunas =
                    Table.ColumnNames(Tabela),

                Colunas =
                    Table.ToColumns(Tabela),

                MapaColunas =

                    Record.FromList(

                        List.Positions(NomesColunas),

                        NomesColunas

                    ),

                EstadoInicial =

                    [

                        Ocorrencias = {}

                    ],

//--------------------------------------------------------------------------
// Validações
//--------------------------------------------------------------------------

                EstadoFinal =

                    List.Accumulate(

                        Validacoes,

                        EstadoInicial,

                        (Estado, Etapa) =>

                            let

                                Posicao =
                                    Record.Field(
                                        MapaColunas,
                                        Etapa[Coluna]
                                    ),

                                Contexto =

                                    [

                                        Coluna = Etapa[Coluna],
                                        Tipo = Etapa[Tipo]

                                    ],

                                OcorrenciasColuna =

                                    fxStgAplicarValidacoesColuna(

                                        Colunas{Posicao},

                                        Etapa[Operadores],

                                        Contexto

                                    )

                            in

                                [

                                    Ocorrencias =

                                        Estado[Ocorrencias] &

                                        {

                                            OcorrenciasColuna

                                        }

                                ]

                    ),

//--------------------------------------------------------------------------
// Consolidação
//--------------------------------------------------------------------------

                QuantidadeLinhas =

                    if List.IsEmpty(Colunas) then

                        0

                    else

                        List.Count(Colunas{0}),

                Ocorrencias =

                    List.Transform(

                        {0 .. QuantidadeLinhas - 1},

                        (i) =>

                            let

                                ListaLinha =

                                    List.Combine(

                                        List.RemoveNulls(

                                            List.Transform(

                                                EstadoFinal[Ocorrencias],

                                                each _{i}

                                            )

                                        )

                                    )

                            in

                                if List.IsEmpty(ListaLinha) then
                                    null
                                else
                                    ListaLinha

                    )

            in

                Table.FromColumns(

                    Colunas &

                    {

                        Ocorrencias

                    },

                    NomesColunas &

                    {

                        "Ocorrencias"

                    }

                )

in

    Resultado;

shared fxStgAplicarValidacoesColuna = (
    valores as list,
    validacoes as list,
    contexto as record
)
as list =>

let

    EstadoInicial =

        List.Repeat(

            { {} },

            List.Count(
                valores
            )

        ),

    EstadoFinal =

        List.Accumulate(

            validacoes,

            EstadoInicial,

            (
                estado,
                operador
            ) =>

                let

                    Funcao =

                        operador[Função],

                    Parametros =

                        Record.FieldOrDefault(

                            operador,

                            "Parametros",

                            null

                        ),

                    ContextoOperador =

                        contexto &

                        [

                            Operador =
                                operador

                        ]

                in

                    List.Transform(

                        List.Positions(
                            valores
                        ),

                        (i) =>

                            let

                                Resultado =

                                    Funcao(

                                        valores{i},

                                        Parametros,

                                        ContextoOperador

                                    ),

                                NovasOcorrencias =

                                    Resultado[Ocorrencias] ?? {}

                            in

                                List.Combine(

                                    {

                                        estado{i},

                                        NovasOcorrencias

                                    }

                                )

                    )

        ),

    Resultado =

        List.Transform(

            EstadoFinal,

            each

                if List.IsEmpty(_) then

                    null

                else

                    _

        )

in

    Resultado;

shared fxSchemaTratamentoNormalizeType = (
    valor as any,
    optional parametros as nullable any
)
as any =>

let

    Parametro =

        if parametros = null then

            null

        else if parametros is list and not List.IsEmpty(parametros) then

            parametros{0}

        else

            parametros,

    Tipo =

        if Parametro = null then
            null
        else
            Record.FieldOrDefault(
                cfgTiposDados,
                Text.Upper(Text.From(Parametro)),
                null
            ),

    Resultado =

        if valor = null then

            null

        else if Tipo = null then

            valor

        else if Tipo = type text then

            try Text.From(valor) otherwise null

        else if Tipo = type number then

            try Number.From(valor) otherwise null

        else if Tipo = type logical then

            try Logical.From(valor) otherwise null

        else if Tipo = type date then

            try Date.From(valor) otherwise null

        else if Tipo = type datetime then

            try DateTime.From(valor) otherwise null

        else if Tipo = type datetimezone then

            try DateTimeZone.From(valor) otherwise null

        else if Tipo = type time then

            try Time.From(valor) otherwise null

        else if Tipo = type duration then

            try Duration.From(valor) otherwise null

        else if Tipo = type binary then

            try Binary.From(valor) otherwise null

        else

            valor

in

    Resultado;

shared fxStgConverterTipoColuna = /*
Responsabilidades:
- Converter uma coluna para um tipo específico.
- Nunca lançar erro.
- Converter valores inválidos para null.
- Preservar null.
*/

(
    valores as list,
    tipo as type
)
as list =>

let

    Conversor =

        if tipo = type text then
            Text.From

        else if tipo = Int64.Type then
            Int64.From

        else if tipo = type number then
            Number.From

        else if tipo = type logical then
            Logical.From

        else if tipo = type date then
            Date.From

        else if tipo = type datetime then
            DateTime.From

        else if tipo = type datetimezone then
            DateTimeZone.From

        else if tipo = type time then
            Time.From

        else if tipo = type duration then
            Duration.From

        else if tipo = type binary then
            Binary.From

        else
            null,

    Resultado =

        if Conversor = null then

            valores

        else

            List.Transform(

                valores,

                each

                    if _ = null then

                        null

                    else

                        try Conversor(_) otherwise null

            )

in

    Resultado
;

shared tstStgAplicarValidacoes = let

    Fonte = srcClientes,

    Pipeline = fxPipeline("tbClientes"),

    Resultado =
        fxNrmAplicarValidacoes(
            Fonte,
            Pipeline
        )
in
    Resultado;

shared fxStgRemoverColunas = /*
Responsabilidades:
- Remover colunas que não pertencem ao Pipeline.
- Nunca adicionar colunas.
- Nunca alterar tipos.
- Nunca ordenar colunas.
*/

(
    Tabela as table,
    Pipeline as record
)
as table =>

let

//--------------------------------------------------------------------------
// Colunas
//--------------------------------------------------------------------------

    Colunas =
        Pipeline[Colunas],

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        if List.IsEmpty(Colunas) then

            #table({}, {})

        else

            Table.SelectColumns(

                Tabela,

                Colunas,

                MissingField.Ignore

            )

in

    Resultado;

shared fxStgOrdenarColunas = /* 
Responsabilidades:
- Ordenar as colunas conforme a definição do Pipeline.
- Nunca adicionar colunas.
- Nunca remover colunas.
- Nunca alterar tipos.
*/

(
    Tabela as table,
    Pipeline as record
)
as table =>

let

//--------------------------------------------------------------------------
// Ordem
//--------------------------------------------------------------------------

    Ordem =
        Pipeline[Ordem],

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        Table.ReorderColumns(

            Tabela,

            Ordem,

            MissingField.Ignore

        )

in

    Resultado;

shared stgPipeline = let

//--------------------------------------------------------------------------
// Defaults
//--------------------------------------------------------------------------

    Defaults =

        [

            COLUNAS =
                [
                Tipo = type any,

                Obrigatório = false,

                Ordem = null,

                Tratamentos =

                    Record.FieldValues(
                        cfgOperadoresTratamentosDefaults
                    ),

                Validações =

                    Record.FieldValues(
                        cfgOperadoresValidacoesDefaults
                    )
                ]
        ],

//--------------------------------------------------------------------------
// Schemas
//--------------------------------------------------------------------------

    Schemas =

        Record.FieldNames(
            cfgSchema
        ),

//--------------------------------------------------------------------------
// Pipeline padrão
//--------------------------------------------------------------------------

    PipelineDefault =

        [

            Schema = "DEFAULTS",

            Pipeline = Defaults

        ],

//--------------------------------------------------------------------------
// Pipelines com Schema
//--------------------------------------------------------------------------

    Pipelines =

        List.Transform(

            Schemas,

            (Schema) =>

                let

                    Definicao =

                        Record.Field(
                            cfgSchema,
                            Schema
                        ),

                    Colunas =

                        Record.FieldNames(
                            Definicao
                        ),

                    Pipeline =

                        Record.FromList(

                            List.Transform(

                                Colunas,

                                (Coluna) =>

                                    Record.Combine({

                                        Defaults,

                                        Record.Field(
                                            Definicao,
                                            Coluna
                                        )

                                    })

                            ),

                            Colunas

                        )

                in

                    [

                        Schema = Schema,

                        Pipeline = Pipeline

                    ]

        ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        Table.FromRecords(

            List.Combine({

                { PipelineDefault },

                Pipelines

            })

        )

in

    Resultado;

shared fxParametroLerParametroPQ = let

    fxParametroLerParametroPowerQuery =

        (
            valor as any
        )

        as nullable record =>

        let

            Metadata =
                Value.Metadata(
                    valor
                ),

            EhParametro =
                Record.HasFields(
                    Metadata,
                    {
                        "IsParameterQuery"
                    }
                )
                and
                Metadata[IsParameterQuery] = true,

            Resultado =

                if not EhParametro then

                    null

                else

                    [

                        Valor =
                            valor,

                        Tipo =
                            fxParametroTipo(
                                Record.FieldOrDefault(
                                    Metadata,
                                    "Type"
                                )
                            ),

                        Obrigatório =
                            Record.FieldOrDefault(
                                Metadata,
                                "IsParameterQueryRequired"
                            ),

                        Permitidos =
                            Record.FieldOrDefault(
                                Metadata,
                                "AllowedValues"
                            )

                    ]

        in

            Resultado

in

    fxParametroLerParametroPowerQuery;

shared stgOperadores = let

//--------------------------------------------------------------------------
// Função de padronização
//--------------------------------------------------------------------------

    fxPadronizar =

        (Tabela as table, Categoria as text) =>

            let

                Renomeada =

                    Table.RenameColumns(

                        Tabela,

                        {

                            {"Name", "Nome"},
                            {"Código", "Codigo"},
                            {"Tipo Entrada", "TipoEntrada"},
                            {"Tipo Saída", "TipoSaida"},
                            {"Severidade Padrão", "SeveridadePadrao"}

                        },

                        MissingField.Ignore

                    ),

                CategoriaAdicionada =

                    Table.AddColumn(

                        Renomeada,

                        "Categoria",

                        each Categoria,

                        type text

                    ),

                Nome =

                    Table.RenameColumns(

                        CategoriaAdicionada,

                        {

                            {"Nome", "Chave"}

                        }

                    ),

                Codigo =

                    Table.RenameColumns(

                        Table.RemoveColumns(
                            CategoriaAdicionada,
                            {"Nome"}
                        ),

                        {

                            {"Codigo", "Chave"}

                        }

                    ),

                CodigoRestaurado =

                    Table.AddColumn(

                        Codigo,

                        "Codigo",

                        each [Chave],

                        type text

                    ),

                Resultado =

                    Table.Combine({

                        Nome,
                        CodigoRestaurado

                    })

            in

                Resultado,

//--------------------------------------------------------------------------
// União
//--------------------------------------------------------------------------

    Fonte =

        Table.Combine({

            fxPadronizar(
                stgParametrosTratamentos,
                "Tratamento"
            ),

            fxPadronizar(
                stgParametrosValidacoes,
                "Validação"
            )

        }),

//--------------------------------------------------------------------------
// Remove duplicidades
//--------------------------------------------------------------------------

    Resultado =

        Table.Distinct(

            Fonte,

            {"Chave"}

        )

in

    Resultado;

shared cfgOperadores = let

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

    Fonte =
        Table.Buffer(
            stgOperadores
        ),

//--------------------------------------------------------------------------
// Operadores
//--------------------------------------------------------------------------

    Operadores =

        List.Transform(

            Table.ToRecords(Fonte),

            each

                Record.AddField(

                    Record.RemoveFields(
                        _,
                        {"Chave"}
                    ),

                    "Parametros",
                    null

                )

        ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        Record.FromList(

            Operadores,

            Fonte[Chave]

        )

in

    Resultado;

shared cfgOperadoresTratamentosDefaults = let

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

    Fonte = 

        Record.SelectFields(

            cfgOperadores,

            List.Select(

                Record.FieldNames(
                    cfgOperadores
                ),

                (Nome) => 
                    let
                        Campo = Record.Field(cfgOperadores, Nome)
                    in
                        Value.Is(Campo, type record) and Campo[Categoria] = "Tratamento" and Campo[Padrão] = true

            )

        )

in

    Fonte;

shared cfgOperadoresValidacoesDefaults = let

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

    Fonte = 

        Record.SelectFields(

            cfgOperadores,

            List.Select(

                Record.FieldNames(
                    cfgOperadores
                ),

                (Nome) => 
                    let
                        Campo = Record.Field(cfgOperadores, Nome)
                    in
                        Value.Is(Campo, type record) and Campo[Categoria] = "Validação" and Campo[Padrão] = true

            )

        )

in

    Fonte;

shared cfgSchema = /*
------------------------------------------------------------------------------
Consulta.: cfgSchema

Descrição.:
Materializa o schema em um Record indexado pelo nome da tabela.

Estrutura:

cfgSchema[Tabela][Coluna]

Ex.:

cfgSchema[tbClientes][CPF]

------------------------------------------------------------------------------
*/

let

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

    Fonte =
        Table.Buffer(
            stgSchema
        ),

//--------------------------------------------------------------------------
// Schema por tabela
//--------------------------------------------------------------------------

    Tabelas =

        Table.Group(

            Fonte,

            {"Tabela"},

            {

                {

                    "Schema",

                    (t) =>

                        Record.FromList(

                            Table.ToRecords(

                                Table.RemoveColumns(

                                    t,

                                    {
                                        "Tabela",
                                        "Coluna"
                                    }

                                )

                            ),

                            t[Coluna]

                        ),

                    type record

                }

            }

        ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        Record.FromList(

            Tabelas[Schema],

            Tabelas[Tabela]

        )
in
    Resultado



/*
let
    Fonte = Table.Buffer(stgSchema),

    Tabelas = Table.Group(
        Fonte,
        {"Tabela"},
        {
            {
                "Schema",
                (t) =>
                    let
                        Colunas = Table.Column(t, "Coluna"),
                        NomesCampos = List.RemoveItems(Table.ColumnNames(t), {"Tabela", "Coluna"}),
                        MatrizDados = Table.ToRows(Table.SelectColumns(t, NomesCampos)),
                        Registros = List.Transform(
                            MatrizDados,
                            (linha) => Record.FromList(linha, NomesCampos)
                        ),
                        Resultado = Record.FromList(Registros, Colunas)
                    in
                        Resultado,
                type record
            }
        }
    ),

    Schema = Record.FromTable(
        Table.RenameColumns(
            Tabelas,
            {
                {"Tabela", "Name"},
                {"Schema", "Value"}
            }
        )
    ),
    tbClientes = Schema[tbClientes],
    CPF = tbClientes[CPF],
    Validações = CPF[Validações]
in
    Validações
*/;

shared cfgPipeline = let

    Resultado =

        Record.TransformFields(

            cfgSchema,

            List.Transform(

                Record.FieldNames(
                    cfgSchema
                ),

                each
                    {
                        _,
                        fxCompilarPipeline
                    }

            )

        )
in
    Resultado;

shared fxCompilarOperadores = /*
------------------------------------------------------------------------------
Função......: fxCompilarOperadores

Descrição...:
Compila uma lista de operadores do Schema em uma estrutura executável.

------------------------------------------------------------------------------
*/

(
    Operadores as nullable list
)
as list =>

let

//--------------------------------------------------------------------------
// Operadores
//--------------------------------------------------------------------------

    ListaOperadores =

        List.RemoveNulls(
            Operadores ?? {}
        ),

//--------------------------------------------------------------------------
// Compilação
//--------------------------------------------------------------------------

    Resultado =

        if List.IsEmpty(ListaOperadores) then

            {}

        else

            List.Buffer(

                List.Transform(

                    ListaOperadores,

                    (Texto) =>

                        let

                            Operador =

                                fxSchemaInterpretarOperador(
                                    Texto
                                ),

                            Definicao =

                                Record.FieldOrDefault(

                                    cfgOperadores,

                                    Operador[Código]

                                )

                        in

                            if Definicao = null then

                                error Error.Record(

                                    "Operador inválido",

                                    "O operador '" &
                                    Operador[Código] &
                                    "' não está cadastrado.",

                                    [
                                        Código = Operador[Código]
                                    ]

                                )

                            else

                                Record.TransformFields(

                                    Definicao,

                                    {
                                        {
                                            "Parametros",
                                            each Operador[Parametros]
                                        }
                                    }

                                )

                )

            )

in

    Resultado;

shared fxSchemaInterpretarOperador = /*
------------------------------------------------------------------------------
Função......: fxSchemaInterpretarOperador

Descrição...:
Interpreta um operador definido no Schema.

Exemplos:

    TRIM
    EMAIL(100)
    BETWEEN(10,20)

Retorno:

    [
        Código = "EMAIL",
        Parametros = {100}
    ]

------------------------------------------------------------------------------
*/

(
    Operador as any
)
as record =>

let

//--------------------------------------------------------------------------
// Texto
//--------------------------------------------------------------------------

    Texto =
        Text.Trim(
            Text.From(Operador)
        ),

//--------------------------------------------------------------------------
// Código
//--------------------------------------------------------------------------

    InicioParametros =
        Text.PositionOf(
            Texto,
            "("
        ),

    Codigo =
        if InicioParametros < 0 then

            Text.Upper(Texto)

        else

            Text.Upper(

                Text.Trim(

                    Text.Start(
                        Texto,
                        InicioParametros
                    )

                )

            ),

//--------------------------------------------------------------------------
// Parâmetros
//--------------------------------------------------------------------------

    ParametrosTexto =
        if InicioParametros < 0 then

            null

        else

            Text.Trim(

                Text.BetweenDelimiters(
                    Texto,
                    "(",
                    ")"
                )

            ),

    Parametros =
        if ParametrosTexto = null or ParametrosTexto = "" then

            null

        else

            List.Buffer(

                List.Transform(

                    List.Select(

                        List.Transform(

                            Text.Split(
                                ParametrosTexto,
                                ","
                            ),

                            each Text.Trim(_)

                        ),

                        each _ <> ""

                    ),

                    each
                        let
                            Numero =
                                try Number.From(_)
                        in
                            if Numero[HasError] then
                                _
                            else
                                Numero[Value]

                )

            )

in

    [

        Código = Codigo,

        Parametros = Parametros

    ];

shared fxCompilarPipelineColuna = /*
------------------------------------------------------------------------------
Função......: fxCompilarPipelineColuna

Descrição...:
Compila uma definição de coluna do cfgSchema para uma estrutura executável.

------------------------------------------------------------------------------
*/

(
    Definicao as record
)
as record =>

let

//--------------------------------------------------------------------------
// Tratamentos
//--------------------------------------------------------------------------

    Tratamentos =

        List.Buffer(

            fxCompilarOperadores(
                Definicao[Tratamentos]
            ) ?? {}

        ),

//--------------------------------------------------------------------------
// Validações
//--------------------------------------------------------------------------

    Validacoes =

        fxCompilarOperadores(
            Definicao[Validações]
        ),

//--------------------------------------------------------------------------
// REQUIRED implícito
//--------------------------------------------------------------------------

    Required =

        if Definicao[Obrigatório] then

            {

                Record.Combine(

                    {

                        cfgOperadores[REQUIRED],

                        [

                            Parametros = null

                        ]

                    }

                )

            }

        else

            {},

//--------------------------------------------------------------------------
// Pipeline de validações
//--------------------------------------------------------------------------

    PipelineValidacoes =

        List.Buffer(

            List.Combine(

                {

                    Required,

                    Validacoes ?? {}

                }

            )

        ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        [

            Tipo = Definicao[Tipo],

            Obrigatório = Definicao[Obrigatório],

            Ativo = Definicao[Ativo],

            Tratamentos = Tratamentos,

            Validações = PipelineValidacoes

        ]

in

    Resultado;

shared fxCompilarPipeline = //==============================================================================
// fxCompilarPipeline
// Compila um Schema em um Pipeline otimizado para execução.
//==============================================================================

(
    Schema as record
)
as record =>

let

//--------------------------------------------------------------------------
// Colunas compiladas
//--------------------------------------------------------------------------

    ColunasCompiladas =

        List.Transform(

            Record.FieldNames(
                Schema
            ),

            (Coluna) =>

                [

                    Coluna = Coluna,

                    Pipeline =
                        fxCompilarPipelineColuna(

                            Record.Field(
                                Schema,
                                Coluna
                            )

                        )

                ]

        ),

    Colunas =

        List.Transform(
            ColunasCompiladas,
            each [Coluna]
        ),

    Ordem =
        Colunas,

//--------------------------------------------------------------------------
// Tipos
//--------------------------------------------------------------------------

    Tipos =

        List.Transform(

            ColunasCompiladas,

            each

                [

                    Coluna = [Coluna],

                    Tipo = [Pipeline][Tipo]

                ]

        ),

//--------------------------------------------------------------------------
// Tratamentos
//--------------------------------------------------------------------------

    Tratamentos =

        List.RemoveNulls(

            List.Transform(

                ColunasCompiladas,

                each

                    if List.IsEmpty(
                        [Pipeline][Tratamentos]
                    ) then

                        null

                    else

                        [

                            Coluna = [Coluna],
                            Tipo = [Pipeline][Tipo],
                            Operadores =
                                [Pipeline][Tratamentos]

                        ]

            )

        ),

//--------------------------------------------------------------------------
// Validações
//--------------------------------------------------------------------------

    Validações =

        List.RemoveNulls(

            List.Transform(

                ColunasCompiladas,

                each

                    if List.IsEmpty(
                        [Pipeline][Validações]
                    ) then

                        null

                    else

                        [

                            Coluna = [Coluna],
                            Tipo = [Pipeline][Tipo],

                            Operadores =
                                [Pipeline][Validações]

                        ]

            )

        ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        [

            Colunas = Colunas,

            Ordem = Ordem,

            Tipos = Tipos,

            Tratamentos = Tratamentos,

            Validações = Validações

        ]

in

    Resultado;

shared fxPipeline = (
    schema as text
)
as record =>

let
    Resultado =
            Record.Field(
                cfgPipeline,
                schema
            )
in
    Resultado;

shared tstPipeline = let
    Fonte = fxPipeline("tbClientes"),
    Validações = Fonte[Validações],
    Validações1 = Validações{3},
    Operadores = Validações1[Operadores]
in
    Operadores;

shared tstCompilarPipeline = let
    Fonte = fxCompilarPipeline(fxSchema("tbClientes")),
    Validações = Fonte[Validações],
    Validações1 = Validações{3}
in
    Validações1;