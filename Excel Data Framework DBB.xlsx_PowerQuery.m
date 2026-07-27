// Power Query from: Excel Data Framework DBB.xlsx
// Pathname: c:\Users\daniel-bighelini\OneDrive\Documentos\Planilhas\Excel Data Framework DBB\Excel Data Framework DBB.xlsx
// Extracted: 2026-07-27T18:07:14.698Z

section Section1;

shared srcTiposDados = /*
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

shared srcOperadores = let
    CategoriaTratamento = "Tratamento",
    CategoriaValidacao = "Validação",
    SeveridadeAviso = "Informação",
    SeveridadeErro = "Erro",
    TabelaParametros = #table(
        type table
        [
            Código = text,
            Descrição = text,
            Função = function,
            TipoEntrada = type,
            TipoSaida = type,
            Padrão = logical,
            Ativo = logical,
            Categoria = text,
            Severidade = text,
            Parâmetros = text
        ],
        {
            {"TRIM", "Remove espaços em branco do início e do final do texto", fxTratamentoTrim, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"UPPER", "Converte todo o texto para letras maiúsculas", fxTratamentoUpper, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"LOWER", "Converte todo o texto para letras minúsculas", fxTratamentoLower, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"PROPER", "Converte a primeira letra de cada palavra para maiúscula", fxTratamentoProper, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"CLEAN", "Remove caracteres não imprimíveis do texto", fxTratamentoClean, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"EMPTYTONULL", "Converte valores vazios ou em branco para null", fxTratamentoEmptyToNull, type any, type any, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"NULLTOEMPTY", "Converte valores nulos (null) em textos vazios", fxTratamentoNullToEmpty, type any, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"SINGLESPACE", "Substitui múltiplos espaços consecutivos por apenas um espaço", fxTratamentoSingleSpace, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"DIGITS", "Mantém apenas os dígitos numéricos do texto", fxTratamentoDigits, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"ALPHANUMERIC", "Mantém apenas letras e números, removendo símbolos e pontuações", fxTratamentoAlphaNumeric, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"ABS", "Retorna o valor absoluto (positivo) de um número", fxTratamentoAbs, type number, type number, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"ROUND", "Arredonda um número decimal para a quantidade de casas especificadas", fxTratamentoRound, type number, type number, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"CPF", "Formata ou extrai apenas os números para o padrão de CPF", fxTratamentoDigits, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"CNPJ", "Formata ou extrai apenas os números para o padrão de CNPJ", fxTratamentoDigits, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"CEP", "Formata ou extrai apenas os números para o padrão de CEP", fxTratamentoDigits, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"NORMALIZEBASIC", "Realiza a normalização básica de textos", fxTratamentoNormalizeBasic, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"NORMALIZETYPE", "Normaliza o tipo de dado conforme o padrão definido", fxTratamentoNormalizeType, type any, type any, false, true, CategoriaTratamento, SeveridadeAviso, null},
            
            {"REQUIRED", "Valida se um campo obrigatório foi preenchido", fxValidacaoREQUIRED, type any, type any, true, true, CategoriaValidacao, SeveridadeErro, null},
            {"CPFVAL", "Valida se o número de CPF informado é matematicamente válido", fxValidacaoCPF, type text, type text, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"CNPJVAL", "Valida se o número de CNPJ informado é matematicamente válido", fxValidacaoCNPJ, type text, type text, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"CEPVAL", "Valida se o formato do CEP informado está correto", fxValidacaoCEP, type text, type text, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"EMAIL", "Valida se a estrutura do endereço de e-mail está correta", fxValidacaoEmail, type text, type text, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"URL", "Valida se a estrutura do endereço web (URL) está correta", fxValidacaoURL, type text, type text, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"LIST", "Valida se o valor pertence a uma lista de opções permitidas", fxValidacaoList, type any, type any, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"DOMAIN", "Valida se o domínio de rede ou e-mail é válido", fxValidacaoDomain, type text, type text, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"SIZE", "Valida se o tamanho ou comprimento do dado está dentro do limite", fxValidacaoSize, type text, type text, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"MIN", "Valida se o valor é maior ou igual ao limite mínimo permitido", fxValidacaoMin, type any, type any, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"MAX", "Valida se o valor é menor ou igual ao limite máximo permitido", fxValidacaoMax, type any, type any, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"INTERVAL", "Valida se o valor está dentro de um intervalo numérico ou temporal específico", fxValidacaoInterval, type any, type any, false, true, CategoriaValidacao, SeveridadeErro, null}
        }
    )
in
    TabelaParametros;

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

shared stgObjetosExcel = 
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

                        fxObjetoIdentificarPeloNome(

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

            Objetos,
            type table [
                Origem = text,
                Nome = text,
                Prefix = nullable text,
                Kind = text,
                Tipo = text,
                Type = type,
                Categoria = text,
                IsStructured = logical
            ]

        )

in

    Resultado;

shared stgObjetosPowerQuery = 
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

                                fxObjetoIdentificarPeloNome(

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

            Objetos,
            type table [
                Origem = text,
                Nome = text,
                Prefix = nullable text,
                Kind = text,
                Tipo = text,
                Type = type,
                Categoria = text,
                IsStructured = logical
            ]
        )

in

    Resultado;

shared stgObjetos = 
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

shared stgTabelasExcel = 
let

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

    Fonte =
        stgObjetosExcel,

    LinhasFiltradas = 
        Table.SelectRows(Fonte, each not Text.Contains([Nome], "!")),

//--------------------------------------------------------------------------
// Metadados
//--------------------------------------------------------------------------

    Registros =

        List.Transform(

            Table.ToRecords(
                LinhasFiltradas
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
            Registros,
            type table [
                Origem = text,
                Nome = text,
                Prefix = nullable text,
                Kind = text,
                Tipo = text,
                Type = type,
                Categoria = text,
                IsStructured = logical,
                Columns = list,
                ColumnCount = number
            ]
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

            Registros,
            type table [
                Origem = text,
                Nome = text,
                Prefix = nullable text,
                Kind = text,
                Tipo = text,
                Type = type,
                Categoria = text,
                IsStructured = logical,
                Columns = list,
                ColumnCount = number
            ]            

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
                    each fxParametroIdentificarTipo(_),
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
                        fxListaNormalizar([Permitidos]),
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
                                fxParametrosLerParametroPQ(
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

;

shared fxParametroIdentificarTipo = (tipo as nullable text) as nullable type =>

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

shared cfgTiposDados = srcTiposDados;

shared cfgTiposObjetos = srcTiposObjetos;

shared srcTiposObjetos = [

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

    ];

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

shared stgOperadores = let
    CategoriaTratamento = "Tratamento",
    CategoriaValidacao = "Validação",

    Operadores =
        srcOperadores,

    Tratamentos =
        Table.Buffer(
            stgParametrosTratamentos
        ),

    Validacoes =
        Table.Buffer(
            stgParametrosValidacoes
        ),

    TratamentosMesclados =
        Table.NestedJoin(
            Operadores,
            {"Código", "Categoria"},
            Tratamentos,
            {"Código", "Categoria"},
            "Tratamentos",
            JoinKind.LeftOuter
        ),

    ValidacoesMesclados =
        Table.NestedJoin(
            TratamentosMesclados,
            {"Código", "Categoria"},
            Validacoes,
            {"Código", "Categoria"},
            "Validações",
            JoinKind.LeftOuter
        ),

    TratamentosExpandidos =
        Table.ExpandTableColumn(
            ValidacoesMesclados,
            "Tratamentos",
            {"Padrão", "Ativo"},
            {"PadrãoTratamento", "AtivoTratamento"}
        ),

    ValidacoesExpandidas =
        Table.ExpandTableColumn(
            TratamentosExpandidos,
            "Validações",
            {"Padrão", "Ativo", "Severidade"},
            {"PadrãoValidação", "AtivoValidação", "SeveridadeValidação"}
        ),

    AtivoAtualizado =
        Table.AddColumn(
            ValidacoesExpandidas,
            "AtivoEfetivo",
            each
                if [Categoria] = CategoriaTratamento and [AtivoTratamento] <> null then
                    [AtivoTratamento]
                else if [Categoria] = CategoriaValidacao and [AtivoValidação] <> null then
                    [AtivoValidação]
                else
                    [Ativo],
            type logical
        ),

    PadraoAtualizado =
        Table.AddColumn(
            AtivoAtualizado,
            "PadrãoEfetivo",
            each
                if [Categoria] = CategoriaTratamento and [PadrãoTratamento] <> null then
                    [PadrãoTratamento]
                else if [Categoria] = CategoriaValidacao and [PadrãoValidação] <> null then
                    [PadrãoValidação]
                else
                    [Padrão],
            type logical
        ),

    SeveridadeAtualizada =
        Table.AddColumn(
            PadraoAtualizado,
            "SeveridadeEfetiva",
            each
                if [Categoria] = CategoriaValidacao and [SeveridadeValidação] <> null then
                    Text.Upper([SeveridadeValidação])
                else
                    Text.Upper([Severidade]),
            type text
        ),

    ColunasRemovidas =
        Table.RemoveColumns(
            SeveridadeAtualizada,
            {
                "Padrão",
                "Ativo",
                "Severidade",
                "PadrãoTratamento",
                "AtivoTratamento",
                "PadrãoValidação",
                "AtivoValidação",
                "SeveridadeValidação"
            }
        ),

    Resultado =
        Table.RenameColumns(
            ColunasRemovidas,
            {
                {"AtivoEfetivo", "Ativo"},
                {"PadrãoEfetivo", "Padrão"},
                {"SeveridadeEfetiva", "Severidade"}
            }
        )

in
    Resultado;

shared stgParametrosTratamentos = let

    Fonte =
        srcParametrosTratamentos,

    LinhasValidas = Table.SelectRows(
        Fonte,
        each [Ativo] = true and
             [Código] <> null and Text.Trim(Text.From([Código])) <> ""
    ),
    
    Tratamentos =
        Table.TransformColumns(
            LinhasValidas,
            {
                {
                    "Código",
                    each if _ = null then
                        null
                    else
                        Text.Upper(
                            Text.Trim(
                                Text.From(_)
                            )
                        ),
                        type text
                }
            }
        ),

    TratamentosTipados =
        Table.TransformColumnTypes(
            Tratamentos,
            {
                {"Padrão", type logical},
                {"Ativo", type logical}
            }
        ),

    TratamentosUnicos =
        Table.Distinct(
            TratamentosTipados,
            {"Código"}
        ),
    Categoria = Table.AddColumn(TratamentosUnicos, "Categoria", each "Tratamento", type text)

in

    Categoria;

shared stgParametrosValidacoes = let

    Fonte =
        srcParametrosValidacoes,

    LinhasValidas = Table.SelectRows(
        Fonte,
        each [Ativo] = true and
             [Código] <> null and Text.Trim(Text.From([Código])) <> ""
    ),
    
    Validacoes =
        Table.TransformColumns(
            LinhasValidas,
            {
                {
                    "Código",
                    each if _ = null then
                        null
                    else
                        Text.Upper(
                            Text.Trim(
                                Text.From(_)
                            )
                        ),
                        type text
                },
                {
                    "Severidade",
                    each if _ = null then
                        null
                    else
                        Text.Upper(
                            Text.Trim(
                                Text.From(_)
                            )
                        ),
                        type text
                }
            }
        ),

    ValidacoesTipados =
        Table.TransformColumnTypes(
            Validacoes,
            {
                {"Padrão", type logical},
                {"Ativo", type logical}
            }
        ),

    ValidacoesUnicos =
        Table.Distinct(
            ValidacoesTipados,
            {"Código"}
        ),
    
    Categoria = Table.AddColumn(ValidacoesUnicos, "Categoria", each "Validação", type text)

in

    Categoria;

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
                    each fxParametroIdentificarTipo(_),
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

                                                fxListaNormalizar(_),

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

                                                fxListaNormalizar(_),

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
;

shared cfgTiposBooleanos = srcTiposBooleanos;

shared srcTiposBooleanos = /*
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

shared cfgConversores = srcConversores;

shared srcConversores = [
    Any = (v as any, local as text) as any => v,
    Text = (v as any, local as text) as any => Text.From(v, local),
    List = (v as any, local as text) as any => fxListaNormalizar(v),
    Int64 = (v as any, local as text) as any => Int64.From(v, local),
    Number = (v as any, local as text) as any => Number.From(v, local),
    Date = (v as any, local as text) as any => Date.From(v, local),
    DateTime = (v as any, local as text) as any => DateTime.From(v, local),
    DateTimeZone = (v as any, local as text) as any => DateTimeZone.From(v, local),
    Time = (v as any, local as text) as any => Time.From(v, local),
    Duration = (v as any, local as text) as any => Duration.From(v),
    Logical = (v as any, local as text) as any => fxParseBooleano(v, local)
];

shared fxListaNormalizar = (
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

shared fxParseBooleano = (valor as any, optional local as nullable text) as nullable logical =>

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
        Table.SelectColumns(stgObjetosPowerQuery, {"Nome", "Categoria", "Tipo"}),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        Table.Sort(

            Objetos,

            {

                {"Categoria", Order.Ascending},

                {"Nome", Order.Ascending}

            }

        )

in

    Resultado;

shared diagTabelasExcel = let

    Fonte =
        Table.SelectColumns(stgTabelasExcel, {"Nome", "Columns"}),

    ColunasTexto =
        Table.TransformColumns(

            Fonte,

            {
                {
                    "Columns",
                    each Text.Combine(_, ";"),
                    type text
                }
            }

        ),

    ColunasRenomeadas = Table.RenameColumns(ColunasTexto,{{"Nome", "Tabela"}, {"Columns", "Colunas"}})

in

    ColunasRenomeadas;

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

shared cfgRESTHeaders = srcRESTHeaders;

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
                cfgCalendarioIntervalos,
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

shared cfgCalendarioIntervalos = // Estrutura utilizada para descobrir o intervalo de datas
// que será utilizado na dimensão de calendario.

let
    cfgIntervalos = {
    // Adicione uma entrada por tabela e sua coluna de data.
    fxCalendarioIntervaloData(srcVendas, "Data")
}
in
    cfgIntervalos;

shared cfgObjetos = 
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

shared fxTratamentoTrim = (
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

shared fxTratamentoUpper = (
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

shared fxTratamentoLower = (
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

shared fxTratamentoProper = (
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

shared fxTratamentoClean = (
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

shared fxTratamentoEmptyToNull = (
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

shared fxTratamentoNullToEmpty = (
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

shared fxTratamentoSingleSpace = (
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

shared fxTratamentoDigits = (
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

shared fxTratamentoAlphaNumeric = (
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

shared fxTratamentoAbs = (
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

shared fxTratamentoRound = (
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
                    "Parâmetros"
                ),

            Mensagem =
                descricao,

            Detalhes =
                detalhes

        ]

in

    Resultado;

shared fxValidacaoList = (
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

shared fxValidacaoDomain = (
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

shared fxValidacaoSize = (
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

shared fxValidacaoMin = (
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

shared fxValidacaoMax = (
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

shared fxValidacaoInterval = (
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

shared fxValidacaoEmail = (
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

shared fxValidacaoURL = (
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

shared fxValidacaoCEP = (
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

shared fxValidacaoCPF = (
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

shared fxValidacaoCNPJ = (
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

shared fxCalendarioIntervaloData = (
    tabela as table,
    coluna as text
) as nullable record =>

let
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

shared fxValidacaoREQUIRED = (
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

shared fxNrmAplicarTratamentos = 
(
    tabela as table,
    pipeline as record
)
as table =>

let

//--------------------------------------------------------------------------
// Tratamentos por coluna
//--------------------------------------------------------------------------

    TratamentosPorColuna =
        Record.FieldOrDefault(
            pipeline,
            "TratamentosPorColuna",
            []
        ),

    TiposPorColuna =
        Record.FieldOrDefault(
            pipeline,
            "TiposPorColuna",
            []
        ),

    Transformacoes =
        List.RemoveNulls(
            List.Transform(
                Table.ColumnNames(tabela),
                each
                    let
                        Coluna = _,
                        Operadores =
                            Record.FieldOrDefault(
                                TratamentosPorColuna,
                                Coluna,
                                null
                            ),
                        TipoPipeline =
                            Record.FieldOrDefault(
                                TiposPorColuna,
                                Coluna,
                                null
                            ),
                        TipoAtual =
                            try Type.TableColumn(Value.Type(tabela), Coluna) otherwise null,
                        TipoFinal =
                            if TipoPipeline = null or TipoPipeline = type any then
                                TipoAtual
                            else
                                TipoPipeline,
                        TratadorCompilado =
                            fxNrmCompilarTratamentosPorColuna(
                                Operadores
                            ),
                        TransformadorFinal =
                            if TipoFinal = null then
                                TratadorCompilado
                            else
                                (valor) =>
                                    let
                                        ResultadoTratamento = TratadorCompilado(valor)
                                    in
                                        if ResultadoTratamento = null then
                                            null
                                        else if Value.Is(ResultadoTratamento, TipoFinal) then
                                            ResultadoTratamento
                                        else
                                            try fxConversor(ResultadoTratamento, TipoFinal) otherwise ResultadoTratamento
                    in
                        if Operadores = null or List.IsEmpty(Operadores) then
                            null
                        else if TipoFinal = null then
                            {
                                Coluna,
                                TransformadorFinal
                            }
                        else
                            {
                                Coluna,
                                TransformadorFinal,
                                TipoFinal
                            }
            )
        ),

    Resultado =
        if List.IsEmpty(Transformacoes) then
            tabela
        else
            Table.TransformColumns(
                tabela,
                Transformacoes,
                null,
                MissingField.Ignore
            )

in

    Resultado;

shared fxNrmExecutarValidacoes = 
(
    Tabela as table,
    Pipeline as record
)
as table =>

let

//--------------------------------------------------------------------------
// Validações
//--------------------------------------------------------------------------

    ValidacoesPorColuna =
        Record.FieldOrDefault(
            Pipeline,
            "ValidacoesPorColuna",
            []
        ),

    TiposPorColuna =
        Record.FieldOrDefault(
            Pipeline,
            "TiposPorColuna",
            []
        ),

    ColunasValidacao =
        Record.FieldNames(ValidacoesPorColuna),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        if List.IsEmpty(ColunasValidacao) then

            Table.AddColumn(
                Tabela,
                "Ocorrencias",
                each null,
                type nullable list
            )

        else

            let

                ColunasValidacao =
                    List.Buffer(
                        Record.FieldNames(ValidacoesPorColuna)
                    ),

                ValidadoresPorColuna =
                    Record.FromList(
                        List.Transform(
                            ColunasValidacao,
                            (Coluna) =>
                                fxNrmCompilarValidacoesPorColuna(
                                    Record.Field(
                                        ValidacoesPorColuna,
                                        Coluna
                                    ),
                                    Record.FieldOrDefault(
                                        TiposPorColuna,
                                        Coluna,
                                        type any
                                    ),
                                    Coluna
                                )
                        ),
                        ColunasValidacao
                    ),

                Resultado =
                    Table.AddColumn(
                        Tabela,
                        "Ocorrencias",
                        each
                            let
                                Linha = _,
                                Ocorrencias =
                                    List.Combine(
                                        List.RemoveNulls(
                                            List.Transform(
                                                ColunasValidacao,
                                                (Coluna) =>
                                                    let
                                                        Validador =
                                                            Record.Field(
                                                                ValidadoresPorColuna,
                                                                Coluna
                                                            ),
                                                        Valor =
                                                            Record.FieldOrDefault(
                                                                Linha,
                                                                Coluna,
                                                                null
                                                            )
                                                    in
                                                        Validador(Valor)
                                            )
                                        )
                                    )
                            in
                                if List.IsEmpty(Ocorrencias) then
                                    null
                                else
                                    Ocorrencias,
                        type nullable list
                    )

            in

                Resultado

in

    Resultado;

shared fxNrmRemoverRegistrosBloqueantes = (
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
    
    RemoverRegistrosBloqueantes =
        fxParametro(
            "Remover_Registros_Bloqueantes"
        ),

//--------------------------------------------------------------------------
// Pipeline
//--------------------------------------------------------------------------

    Tratada =

        if ProcessarTratamentos then

            fxNrmAplicarTratamentos(
                Tabela,
                Pipeline
            )

        else

            Tabela,

    Validada =

        if ProcessarValidacoes then

            fxNrmExecutarValidacoes(
                Tratada,
                Pipeline
            )

        else

            Tratada,

    Resultado =

        if RemoverRegistrosBloqueantes then
            fxNrmRemoverRegistrosBloqueantes(
                Validada
            )
        
        else
            
            Validada

in

    Resultado;

shared parTabelaCategoriasConsultasPQ = "tbSobreCategoriasConsultasPQ" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared fxStgPreparar = 
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

shared fxStgGarantirColunas = 
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

shared fxStgRemoverColunas = 
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

shared fxStgAplicarTipos = (
    Tabela as table,
    Pipeline as record,
    optional Identificar_Tipos as nullable logical,
    optional Amostra as nullable number
)
as table =>

let

//--------------------------------------------------------------------------
// Configuração
//--------------------------------------------------------------------------

    IdentificarTipos =
        if Identificar_Tipos = null then
            true
        else
            Identificar_Tipos,

//--------------------------------------------------------------------------
// Tipos
//--------------------------------------------------------------------------

    Tipos =
        Pipeline[Tipos],

//--------------------------------------------------------------------------
// Tipos efetivos
//--------------------------------------------------------------------------

    TiposEfetivos =

        if
            List.IsEmpty(Tipos)
            or not IdentificarTipos
        then

            []

        else

            fxStgIdentificarTiposColunas(

                Tabela,

                List.Transform(

                    Tipos,

                    each [Coluna]

                ),

                Amostra

            ),

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

                                    if
                                        TipoConfigurado <> type any
                                        or not IdentificarTipos
                                    then

                                        TipoConfigurado

                                    else

                                        Record.Field(
                                            TiposEfetivos,
                                            Coluna
                                        )

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

shared fxStgOrdenarColunas = (
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

    OrdemAtual =
        Table.ColumnNames(
            Tabela
        ),

    Reordenar =
        OrdemAtual <> Ordem,

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        if Reordenar then

            Table.ReorderColumns(

                Tabela,

                Ordem,

                MissingField.Ignore

            )

        else

            Tabela

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
// Configuração
//--------------------------------------------------------------------------

    IdentificarTiposColunasAny =
        fxParametro(
            "Identificar_Tipos_Colunas"
        ),


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
                Pipeline,
                IdentificarTiposColunasAny
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

shared fxTratamentoNormalizeBasic = 
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

shared fxObjetoIdentificarTipo = 
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

shared fxObjetoIdentificarPeloValor = 
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
        fxObjetoIdentificarTipo(
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

shared fxObjetoIdentificarPeloNome = 
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

shared fxParametrosIdentificarParametroPQ = 
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

shared fxNrmCompilarTratamentosPorColuna = (
    Operadores as list
)
as function =>

let
    OperadoresBuffer =
        List.Buffer(
            Operadores ?? {}
        )

in
    if List.IsEmpty(OperadoresBuffer) then
        (valor) => valor
    else
        (valor) =>
            List.Accumulate(
                OperadoresBuffer,
                valor,
                (Estado, Operador) =>
                    Operador[Função](
                        Estado,
                        Record.FieldOrDefault(
                            Operador,
                            "Parâmetros",
                            null
                        )
                    )
            );

shared fxNrmCompilarValidacoesPorColuna = (
    Operadores as list,
    Tipo as type,
    Coluna as text
)
as function =>

let
    OperadoresBuffer =
        List.Buffer(
            Operadores ?? {}
        )

in
    if List.IsEmpty(OperadoresBuffer) then
        (valor) => null
    else
        (valor) =>
            let
                Ocorrencias =
                    List.Accumulate(
                        OperadoresBuffer,
                        {},
                        (Estado, Operador) =>
                            let
                                Resultado =
                                    Operador[Função](
                                        valor,
                                        Record.FieldOrDefault(
                                            Operador,
                                            "Parâmetros",
                                            null
                                        ),
                                        [
                                            Coluna = Coluna,
                                            Tipo = Tipo,
                                            Operador = Operador
                                        ]
                                    ),
                                NovasOcorrencias =
                                    Resultado[Ocorrencias] ?? {}
                            in
                                Estado & NovasOcorrencias
                    )
            in
                if List.IsEmpty(Ocorrencias) then
                    null
                else
                    Ocorrencias;

shared fxStgIdentificarTiposColunas = (
    tabela as table,
    colunas as list,
    optional amostra as nullable number
)
as record =>

let

//--------------------------------------------------------------------------
// Configuração
//--------------------------------------------------------------------------

    QuantidadeAmostra =
        if amostra = null then
            20
        else
            Number.RoundDown(amostra),

    TipoTabela =
        Value.Type(tabela),

//--------------------------------------------------------------------------
// Tipos declarados
//--------------------------------------------------------------------------

    TiposDeclarados =

        Record.FromList(

            List.Transform(

                colunas,

                each
                    Type.TableColumn(
                        TipoTabela,
                        _
                    )

            ),

            colunas

        ),

//--------------------------------------------------------------------------
// Colunas que precisam ser inferidas
//--------------------------------------------------------------------------

    ColunasAny =

        List.Select(

            colunas,

            each Record.Field(
                TiposDeclarados,
                _
            ) = type any

        ),

//--------------------------------------------------------------------------
// Inferência
//--------------------------------------------------------------------------

    TiposInferidos =

        if List.IsEmpty(ColunasAny) then

            []

        else

            let

                Listas =

                    Table.ToColumns(

                        Table.SelectColumns(

                            tabela,
                            ColunasAny,
                            MissingField.Ignore

                        )

                    )

            in

                Record.FromList(

                    List.Transform(

                        Listas,

                        (Lista) =>

                            let

                                Valores =

                                    List.FirstN(

                                        List.RemoveNulls(
                                            Lista
                                        ),

                                        QuantidadeAmostra

                                    ),

                                Tipos =

                                    List.Distinct(

                                        List.Transform(

                                            Valores,

                                            Value.Type

                                        )

                                    )

                            in

                                if List.Count(Tipos) = 1 then
                                    Tipos{0}
                                else
                                    type any

                    ),

                    ColunasAny

                ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        Record.Combine({

            TiposDeclarados,
            TiposInferidos

        })

in

    Resultado;

shared tstPadronizar = let

    Fonte = srcClientes,

    Pipeline =
            fxPipeline("tbClientes"),

    Resultado =
        fxNrmAplicarTratamentos(
            Fonte,
            Pipeline
        )
in
    Resultado;

shared tstAplicarTipos = let

    Fonte =
        srcVendas,

    Pipeline =
        fxPipeline("tbVendas"),

    Resultado =
        fxStgAplicarTipos(
            Fonte,
            Pipeline
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

shared fxTratamentoNormalizeType = (
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

shared fxParametrosLerParametroPQ = let

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
                            fxParametroIdentificarTipo(
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
            each Record.RemoveFields(_, {"Código"})
        ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =
        Record.FromList(
            Operadores,
            Fonte[Código]
        )

in

    Resultado;

shared cfgSchema = 
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
;

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
                        fxPìpelineCompilar
                    }

            )

        )
in
    Resultado;

shared fxPipelineInterpretarOperador = 
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

        Parâmetros = Parametros

    ];

shared fxPipelineCompilarOperadores = (
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

                                fxPipelineInterpretarOperador(
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
                                            "Parâmetros",
                                            each Operador[Parâmetros]
                                        }
                                    }

                                )

                )

            )

in

    Resultado;

shared fxPipelineCompilarColuna = 
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

            fxPipelineCompilarOperadores(
                Definicao[Tratamentos]
            ) ?? {}

        ),

//--------------------------------------------------------------------------
// Validações
//--------------------------------------------------------------------------

    Validacoes =

        fxPipelineCompilarOperadores(
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

                            Parâmetros = null

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

shared fxPìpelineCompilar = 
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
                        fxPipelineCompilarColuna(

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

    TiposPorColuna =
        Record.FromList(
            List.Transform(
                Tipos,
                each [Tipo]
            ),
            List.Transform(
                Tipos,
                each [Coluna]
            )
        ),

    TratamentosPorColuna =
        Record.FromList(
            List.Transform(
                Tratamentos,
                each [Operadores]
            ),
            List.Transform(
                Tratamentos,
                each [Coluna]
            )
        ),

    ValidacoesPorColuna =
        Record.FromList(
            List.Transform(
                Validações,
                each [Operadores]
            ),
            List.Transform(
                Validações,
                each [Coluna]
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

            TiposPorColuna = TiposPorColuna,

            Tratamentos = Tratamentos,

            TratamentosPorColuna = TratamentosPorColuna,

            Validações = Validações,

            ValidacoesPorColuna = ValidacoesPorColuna

        ]

in

    Resultado
;

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
    Fonte = fxPipeline("tbClientes")
in
    Fonte;

shared tstCompilarPipeline = let
    Fonte = fxPìpelineCompilar(fxSchema("tbClientes")),
    Tratamentos = Fonte[Tratamentos],
    Tratamentos1 = Tratamentos{0},
    Operadores = Tratamentos1[Operadores],
    Operadores1 = Operadores{0}
in
    Operadores1;

shared tstnrmAplicarTratamentos = let

    Fonte = srcVendas,

    Pipeline = fxPipeline("tbVendas"),

    Resultado =
        fxNrmAplicarTratamentos(
            Fonte,
            Pipeline
        )
in
    Resultado;

shared tstNrmExecutarValidacoes = let

    Fonte = srcClientes,

    Pipeline = fxPipeline("tbClientes"),

    Tratamentos =
        fxNrmAplicarTratamentos(
            Fonte,
            Pipeline
        ),

    Validacoes =
        fxNrmExecutarValidacoes(
            Tratamentos,
            Pipeline
        )
in
    Validacoes;

shared srcRESTHeaders = /*
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