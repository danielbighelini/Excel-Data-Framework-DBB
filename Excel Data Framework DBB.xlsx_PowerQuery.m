// Power Query from: Excel Data Framework DBB.xlsx
// Pathname: c:\Users\daniel-bighelini\OneDrive\Documentos\Planilhas\Excel Data Framework DBB\Excel Data Framework DBB.xlsx
// Extracted: 2026-07-29T19:37:43.388Z

section Section1;

//==============================================================================
// TIPOS DE DADOS - Configuração Central
//==============================================================================
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

    Resultado;
shared srcCategoriasPowerQuery = let
    Fonte = srcWorkbook{[Name=parTabelaCategoriasConsultasPQ]}[Content]
in
    Fonte;

shared srcWorkbook = let
    srcWorkbook = Table.Buffer(Excel.CurrentWorkbook())
in
    srcWorkbook;
shared srcParametrosExcel = let
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
            // Tratamentos
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
            {"REPLACE", "Substitui todas as ocorrências de um texto por outro.", fxTratamentoReplace, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"LEFT", "Mantém os N primeiros caracteres do texto.", fxTratamentoLeft, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"RIGHT", "Mantém os N últimos caracteres do texto.", fxTratamentoRight, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"MID", "Extrai uma quantidade de caracteres a partir de uma posição específica.", fxTratamentoMid, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"BEFORE", "Extrai o texto localizado antes de um delimitador informado.", fxTratamentoBefore, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"AFTER", "Extrai o texto localizado após um delimitador informado.", fxTratamentoAfter, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"PREFIX", "Adiciona um prefixo ao início do texto.", fxTratamentoPrefix, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"SUFFIX", "Adiciona um sufixo ao final do texto.", fxTratamentoSuffix, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"PADLEFT", "Completa o texto à esquerda até atingir o comprimento especificado.", fxTratamentoPadLeft, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"PADRIGHT", "Completa o texto à direita até atingir o comprimento especificado.", fxTratamentoPadRight, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"REMOVECHARS", "Remove todos os caracteres pertencentes a uma lista informada.", fxTratamentoRemoveChars, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"KEEPCHARS", "Mantém apenas os caracteres pertencentes a uma lista informada.", fxTratamentoKeepChars, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"REMOVEACCENTS", "Remove acentos e caracteres diacríticos do texto.", fxTratamentoRemoveAccents, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"PUNCTUATION", "Remove todos os sinais de pontuação do texto.", fxTratamentoPunctuation, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},

            // Validações
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

//==============================================================================
// PARÂMETROS - Configuração do Framework
//==============================================================================

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

shared srcParametrosCalendario = let
    Fonte = srcWorkbook{[Name=parTabelaParametrosCalendario]}[Content]
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

shared cfgPrefixosPowerQuery = List.Buffer(
    Record.FieldNames(cfgCategoriasPowerQuery)
);

shared fxParametro = (
    parametro as text,
    optional valorPadrao as nullable any
) as any =>

let
    Parametros = cfgParametros,

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

shared stgObjetosPowerQuery = let

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

    Fonte =

        srcObjetosPowerQuery,

//--------------------------------------------------------------------------
// Objetos
//--------------------------------------------------------------------------

    Objetos =
        List.RemoveNulls(
            List.Transform(
                Fonte,
                (Nome) =>
                    let
                        Metadados =
                            try
                                fxObjetoIdentificarPeloNome(
                                    Nome,
                                    "PowerQuery"
                                )
                            otherwise
                                null
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
                Kind = nullable text,
                Tipo = nullable text,
                Type = nullable type,
                Categoria = nullable text,
                IsStructured = nullable logical
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

//--------------------------------------------------------------------------
// Fonte
//--------------------------------------------------------------------------

    Fonte =
        #sections[Section1],

//--------------------------------------------------------------------------
// Objetos
//--------------------------------------------------------------------------

    Objetos =
        Record.FieldNames(Fonte),

//--------------------------------------------------------------------------
// Prefixo dos parâmetros
//--------------------------------------------------------------------------

    PrefixoParametro =

        List.First(

            List.Select(

                cfgPrefixosPowerQuery,

                each
                    Record.Field(
                        cfgCategoriasPowerQuery,
                        _
                    )[Categoria] = "Parâmetros"

            ),

            null

        ),

//--------------------------------------------------------------------------
// Nomes candidatos
//--------------------------------------------------------------------------

    NomesCandidatos =

        List.Select(

            Objetos,

            each Text.StartsWith(
                _,
                PrefixoParametro
            )

        ),

//--------------------------------------------------------------------------
// Parâmetros
//--------------------------------------------------------------------------

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
        if tipo = null then null
        else
            Text.Upper(
                Text.Trim(
                    Text.From(tipo)
                )
            )
in
    if Nome = null then null
    else if Record.HasFields(cfgTiposDados, Nome) then
        Record.Field(cfgTiposDados, Nome)
    else
        error Error.Record(
            "Tipo inválido",
            Text.Format(
                "O tipo '#{0}' não está cadastrado.",
                {tipo}
            ),
            [
                Tipo = tipo,
                TiposDisponiveis =
                    Text.Combine(
                        List.Sort(
                            Record.FieldNames(cfgTiposDados)
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
        each 
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
        each 
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

shared stgParametrosCalendario = let
    Fonte =
        srcParametrosCalendario,

    Tipos =
        Table.TransformColumnTypes(
            Fonte,
            {
                {"Nome da Coluna", type text},
                {"Código", type text},
                {"Ordem", Int64.Type},
                {"Ativo", type logical}
            }
        ),

    Trim =
        Table.TransformColumns(
            Tipos,
            {
                {"Nome da Coluna", each Text.Trim(_), type text},
                {"Código", each Text.Upper(Text.Trim(_)), type text}
            }
        ),

    FiltrarAtivos =
        Table.SelectRows(
            Trim,
            each [Ativo]
        ),

    FiltrarValidos =
        Table.SelectRows(
            FiltrarAtivos,
            each
                [Código] <> null and
                [Código] <> "" and
                [Nome da Coluna] <> null and
                [Nome da Coluna] <> ""
        ),

    RemoverDuplicados =
        Table.Distinct(
            FiltrarValidos,
            {"Código"}
        ),

    Ordenar =
        Table.Sort(
            RemoverDuplicados,
            {
                {"Ordem", Order.Ascending},
                {"Nome da Coluna", Order.Ascending}
            }
        ),

    Buffer =
        Table.Buffer(Ordenar)

in
    Buffer;

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

shared stgDados = let
    Fonte =
        fxOrigem(),

    Tabela =
        if
            Value.Is(Fonte, type table)
            and Table.HasColumns(Fonte, "Dados")
            and not Table.IsEmpty(Fonte)
            and Value.Is(Fonte{0}[Dados], type table)
        then
            Table.Combine(Fonte[Dados])
        else
            fxOrigemComoTabela(Fonte),

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
    Separador = if separador = null then ";" else separador,
    RemoverVazios = if removerVazios = null then true else removerVazios,
    Trim = if trim = null then true else trim,

    Lista =
        if valor = null then null
        else if Value.Is(valor, type list) then valor
        else Text.Split(Text.From(valor), Separador),

    ListaTratada =
        if Lista = null then null
        else
            List.Transform(
                Lista,
                each if Trim then Text.Trim(Text.From(_)) else Text.From(_)
            ),

    Resultado =
        if ListaTratada = null then null
        else if RemoverVazios then
            List.RemoveItems(ListaTratada, {""})
        else
            ListaTratada

in
    if Resultado = null then null 
    else List.Buffer(List.Distinct(Resultado));
shared fxParseBooleano = (valor as any, optional local as nullable text) as nullable logical =>

let
    Local = if local = null then "pt-BR" else local,

    Nome =
        if valor = null then null
        else
            Text.Upper(
                Text.Trim(
                    Text.From(valor, Local)
                )
            )
in
    if Nome = null then null
    else if Record.HasFields(cfgTiposBooleanos, Nome) then
        Record.Field(cfgTiposBooleanos, Nome)
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
shared fxOrigemComoTabela = (Valor as any) as table =>

let
    Resultado =
        if Valor = null then
            #table(
                type table [Value = any],
                {}
            )
        else if Value.Is(Valor, type table) then
            Valor
        else if Value.Is(Valor, type record) then
            Record.ToTable(Valor)
        else if Value.Is(Valor, type list) then
            Table.FromList(
                Valor,
                Splitter.SplitByNothing(),
                {"Value"}
            )
        else
            #table(
                type table [Value = any],
                {{Valor}}
            )

in
    Resultado;

shared fxConversor = (valor as any, tipo as nullable type, optional local as nullable text) as any =>
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
    Fonte = srcClientes,
    Preparada = fxStgAplicar(Fonte, parTabelaClientes),
    Resultado = Table.Distinct(Preparada)
in
    Resultado;

shared stgProdutos = let
    Fonte = srcProdutos,
    Preparada = fxStgAplicar(Fonte, parTabelaProdutos),
    Resultado = Table.Distinct(Preparada)
in
    Resultado;

shared stgVendas = let
    Fonte = srcVendas,
    Preparada = fxStgAplicar(Fonte, parTabelaVendas),
    Resultado = Table.Distinct(Preparada)
in
    Resultado;

shared nrmClientes = let
    Fonte = qaClientes,
    Valida = fxQaFiltrarPorStatus(Fonte, "OK"),
    Normalizada = fxNrmAplicar(Valida, parTabelaClientes),
    RegistrosUnicos = Table.Distinct(Normalizada, {"CPF"})
in
    RegistrosUnicos;

shared dimClientes = let
    Fonte = nrmClientes,
    Chaves = Table.AddIndexColumn(Fonte, "IDCliente", 1, 1, Int64.Type),
    Reordenada = Table.ReorderColumns(Chaves, 
        {"IDCliente", "CPF", "Nome", "DataNascimento", "Cidade", "Estado"})
in
    Table.Buffer(Reordenada);

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
    Fonte =
        Csv.Document(
            Content,
            [
                Delimiter = ";",
                Encoding = 65001,
                QuoteStyle = QuoteStyle.Csv
            ]
        ),

    Cabecalhos =
        if Table.IsEmpty(Fonte) then
            Fonte
        else
            Table.PromoteHeaders(
                Fonte,
                [PromoteAllScalars = true]
            ),

    SemLinhasVazias =
        Table.SelectRows(
            Cabecalhos,
            each
                List.AnyTrue(
                    List.Transform(
                        Record.FieldValues(_),
                        (Valor) =>
                            Valor <> null
                            and Text.Trim(Text.From(Valor)) <> ""
                    )
                )
        )

in
    SemLinhasVazias;

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

shared cfgCalendarioAtributos = let
    Definicoes =
    [
        DATA = [
            Tipo = type date,
            Categoria = "Base",
            Funcao = (Data as date, Cultura as text) => Data
        ],

        ANO = [
            Tipo = Int64.Type,
            Categoria = "Ano",
            Funcao = (Data as date, Cultura as text) => Date.Year(Data)
        ],

        SEMESTRE = [
            Tipo = Int64.Type,
            Categoria = "Ano",
            Funcao = (Data as date, Cultura as text) => Number.RoundUp(Date.Month(Data) / 6)
        ],

        TRIMESTRE = [
            Tipo = Int64.Type,
            Categoria = "Ano",
            Funcao = (Data as date, Cultura as text) => Date.QuarterOfYear(Data)
        ],

        MES = [
            Tipo = Int64.Type,
            Categoria = "Mês",
            Funcao = (Data as date, Cultura as text) => Date.Month(Data)
        ],

        NOME_MES = [
            Tipo = type text,
            Categoria = "Mês",
            Funcao = (Data as date, Cultura as text) => Date.MonthName(Data, Cultura)
        ],

        MES_ABREVIADO = [
            Tipo = type text,
            Categoria = "Mês",
            Funcao = (Data as date, Cultura as text) => Date.ToText(Data, "MMM", Cultura)
        ],

        MES_ANO = [
            Tipo = type text,
            Categoria = "Mês",
            Funcao = (Data as date, Cultura as text) => Date.ToText(Data, "MM/yyyy", Cultura)
        ],

        ANO_MES = [
            Tipo = Int64.Type,
            Categoria = "Mês",
            Funcao = (Data as date, Cultura as text) => Date.Year(Data) * 100 + Date.Month(Data)
        ],

        ANO_TRIMESTRE = [
            Tipo = type text,
            Categoria = "Ano",
            Funcao = (Data as date, Cultura as text) =>
                Text.From(Date.Year(Data)) & "/T" & Text.From(Date.QuarterOfYear(Data))
        ],

        DIA = [
            Tipo = Int64.Type,
            Categoria = "Dia",
            Funcao = (Data as date, Cultura as text) => Date.Day(Data)
        ],

        DIA_ANO = [
            Tipo = Int64.Type,
            Categoria = "Dia",
            Funcao = (Data as date, Cultura as text) => Date.DayOfYear(Data)
        ],

        SEMANA_ANO = [
            Tipo = Int64.Type,
            Categoria = "Semana",
            Funcao = (Data as date, Cultura as text) => Date.WeekOfYear(Data)
        ],

        SEMANA_MES = [
            Tipo = Int64.Type,
            Categoria = "Semana",
            Funcao = (Data as date, Cultura as text) => Number.IntegerDivide(Date.Day(Data) - 1, 7) + 1
        ],

        DIA_SEMANA = [
            Tipo = Int64.Type,
            Categoria = "Semana",
            Funcao = (Data as date, Cultura as text) => Date.DayOfWeek(Data) + 1
        ],

        NOME_DIA_SEMANA = [
            Tipo = type text,
            Categoria = "Semana",
            Funcao = (Data as date, Cultura as text) => Date.DayOfWeekName(Data, Cultura)
        ],

        DIA_SEMANA_ABREVIADO = [
            Tipo = type text,
            Categoria = "Semana",
            Funcao = (Data as date, Cultura as text) => Date.ToText(Data, "ddd", Cultura)
        ],

        FINAL_DE_SEMANA = [
            Tipo = type logical,
            Categoria = "Semana",
            Funcao = (Data as date, Cultura as text) => List.Contains({0, 6}, Date.DayOfWeek(Data))
        ],

        DIA_UTIL = [
            Tipo = type logical,
            Categoria = "Semana",
            Funcao = (Data as date, Cultura as text) => not List.Contains({0, 6}, Date.DayOfWeek(Data))
        ],

        BIMESTRE = [
            Tipo = Int64.Type,
            Categoria = "Ano",
            Funcao = (Data as date, Cultura as text) => Number.RoundUp(Date.Month(Data) / 2)
        ],

        QUADRIMESTRE = [
            Tipo = Int64.Type,
            Categoria = "Ano",
            Funcao = (Data as date, Cultura as text) => Number.RoundUp(Date.Month(Data) / 4)
        ],

        DECADA = [
            Tipo = Int64.Type,
            Categoria = "Ano",
            Funcao = (Data as date, Cultura as text) => Number.IntegerDivide(Date.Year(Data), 10) * 10
        ],

        SECULO = [
            Tipo = Int64.Type,
            Categoria = "Ano",
            Funcao = (Data as date, Cultura as text) => Number.RoundUp(Date.Year(Data) / 100)
        ],

        INICIO_MES = [
            Tipo = type date,
            Categoria = "Datas",
            Funcao = (Data as date, Cultura as text) => Date.StartOfMonth(Data)
        ],

        FIM_MES = [
            Tipo = type date,
            Categoria = "Datas",
            Funcao = (Data as date, Cultura as text) => Date.EndOfMonth(Data)
        ],

        INICIO_TRIMESTRE = [
            Tipo = type date,
            Categoria = "Datas",
            Funcao = (Data as date, Cultura as text) => Date.StartOfQuarter(Data)
        ],

        FIM_TRIMESTRE = [
            Tipo = type date,
            Categoria = "Datas",
            Funcao = (Data as date, Cultura as text) => Date.EndOfQuarter(Data)
        ],

        INICIO_ANO = [
            Tipo = type date,
            Categoria = "Datas",
            Funcao = (Data as date, Cultura as text) => Date.StartOfYear(Data)
        ],

        FIM_ANO = [
            Tipo = type date,
            Categoria = "Datas",
            Funcao = (Data as date, Cultura as text) => Date.EndOfYear(Data)
        ],

        ANO_FISCAL = [
            Tipo = Int64.Type,
            Categoria = "Fiscal",
            Funcao = (Data as date, Cultura as text) => Date.Year(Data)
        ],

        TRIMESTRE_FISCAL = [
            Tipo = Int64.Type,
            Categoria = "Fiscal",
            Funcao = (Data as date, Cultura as text) => Date.QuarterOfYear(Data)
        ],

        MES_FISCAL = [
            Tipo = Int64.Type,
            Categoria = "Fiscal",
            Funcao = (Data as date, Cultura as text) => Date.Month(Data)
        ],

        SEMANA_FISCAL = [
            Tipo = Int64.Type,
            Categoria = "Fiscal",
            Funcao = (Data as date, Cultura as text) => Date.WeekOfYear(Data)
        ],

        FERIADO = [
            Tipo = type logical,
            Categoria = "Feriados",
            Funcao = (Data as date, Cultura as text) => false
        ],

        NOME_FERIADO = [
            Tipo = type text,
            Categoria = "Feriados",
            Funcao = (Data as date, Cultura as text) => null
        ],

        TIMESTAMP = [
            Tipo = type datetime,
            Categoria = "Base",
            Funcao = (Data as date, Cultura as text) => DateTime.From(Data)
        ]
    ],

    Tabela =
        Table.AddColumn(
            stgParametrosCalendario,
            "Atributo",
            each Record.Field(Definicoes, [Código]),
            type record
        ),

    Expandido =
        Table.ExpandRecordColumn(
            Tabela,
            "Atributo",
            {"Tipo", "Categoria", "Funcao"}
        ),

    Registro =
        Record.FromList(
            List.Transform(
                Table.ToRecords(Expandido),
                each [
                    Nome = [Nome da Coluna],
                    Ordem = [Ordem],
                    Tipo = [Tipo],
                    Categoria = [Categoria],
                    Funcao = [Funcao]
                ]
            ),
            Expandido[Código]
        )
in
    Registro;

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

shared parTabelaParametrosCalendario = "tbParametrosCalendario" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];
shared parTabelaSchema = "tbSchema" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];
shared parTabelaClientes = "tbClientes" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];
shared parTabelaProdutos = "tbProdutos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];
shared parTabelaVendas = "tbVendas" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];
shared parTabelaDadosGenericos = "tbDadosGenericos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared nrmProdutos = let
    Fonte = qaProdutos,
    Valida = fxQaFiltrarPorStatus(Fonte, "OK"),
    Normalizada = fxNrmAplicar(Valida, parTabelaProdutos),
    RegistrosUnicos = Table.Distinct(Normalizada, {"Código"})
in
    RegistrosUnicos;

shared nrmVendas = let
    Fonte = qaVendas,
    Valida = fxQaFiltrarPorStatus(Fonte, "OK"),
    Normalizada = fxNrmAplicar(Valida, parTabelaVendas),
    ValorTotal = Table.AddColumn(Normalizada, "ValorTotal", each [Quantidade] * [ValorUnitário], type number)
in
    ValorTotal;

shared nrmDados = let
    Fonte = stgDados
in
    Fonte;

shared dimProdutos = let
    Fonte = nrmProdutos,
    Chaves = Table.AddIndexColumn(Fonte, "IDProduto", 1, 1, Int64.Type),
    Reordenada = Table.ReorderColumns(Chaves, 
        {"IDProduto", "Código", "Descrição", "Categoria", "PreçoLista"})
in
    Table.Buffer(Reordenada);

shared fxCalendarioBase = (
    DataInicial as date,
    DataFinal as date,
    optional PrimeiroDiaSemana as nullable number
)
as table =>

let
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
            { Datas },
            type table [
                Data = date
            ]
        )

in
    Calendario;

shared fxCalendario = (
    Calendario as table,
    optional Cultura as nullable text
)
as table =>

let
    Idioma =
        if Cultura = null then
            "pt-BR"
        else
            Cultura,

    Atributos =
        List.Buffer(
            Table.ToRecords(
                Record.ToTable(cfgCalendarioAtributos)
            )
        ),

    Resultado =
        List.Accumulate(

            Atributos,

            Calendario,

            (Tabela, Atual) =>

                let
                    Definicao = Atual[Value]
                in
                    Table.AddColumn(
                        Tabela,
                        Definicao[Nome],
                        each Definicao[Funcao]([Data], Idioma),
                        Definicao[Tipo]
                    )

        )
in
    
    Resultado
;

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
shared fxTratamentoTrim = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else Text.Trim(Text.From(valor));
shared fxTratamentoUpper = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else Text.Upper(Text.From(valor));
shared fxTratamentoLower = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else Text.Lower(Text.From(valor));
shared fxTratamentoProper = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else Text.Proper(Text.From(valor));
shared fxTratamentoClean = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else Text.Clean(Text.From(valor));
shared fxTratamentoEmptyToNull = (valor as any, optional parametros as nullable any) as any =>
    if valor = "" then null else valor;
shared fxTratamentoNullToEmpty = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then "" else valor;
shared fxTratamentoSingleSpace = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null
    else
        Text.Combine(
            List.Select(
                Text.SplitAny(
                    Text.Trim(Text.From(valor)),
                    " "
                ),
                each _ <> ""
            ),
            " "
        );
shared fxTratamentoDigits = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null 
    else Text.Select(Text.From(valor), {"0".."9"});
shared fxTratamentoAlphaNumeric = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null 
    else Text.Select(Text.From(valor), {"A".."Z","a".."z","0".."9"});
shared fxTratamentoAbs = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null 
    else Number.Abs(Number.From(valor));
shared fxTratamentoRound = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else
        Number.Round(Number.From(valor), Number.From(parametros));

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
        Record.FieldOrDefault(contexto, "Operador", []),

    Resultado =
        [
            Codigo = codigo,
            Severidade = Record.FieldOrDefault(Operador, "Severidade"),
            Coluna = Record.FieldOrDefault(contexto, "Coluna"),
            Valor = valor,
            Tipo = Record.FieldOrDefault(contexto, "Tipo"),
            Parametros = Record.FieldOrDefault(Operador, "Parâmetros"),
            Mensagem = descricao,
            Detalhes = detalhes
        ]

in
    Resultado;

//==============================================================================
// TRATAMENTOS - Treatment Functions
// Documentação completa em: https://github.com/seu-usuario/seu-framework
//==============================================================================
shared fxValidacaoList = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let
    Lista = if parametros = null then {} else parametros,

    Valido =
        if valor = null then true
        else List.Contains(Lista, Text.From(valor)),

    Ocorrencias =
        if Valido then null
        else
            {
                fxSchemaOcorrencia(
                    "LIST",
                    contexto,
                    valor,
                    valor,
                    "Valor não pertence à lista permitida.",
                    [Permitidos = Lista]
                )
            }

in
    [Valor = valor, Ocorrencias = Ocorrencias];
shared fxValidacaoDomain = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let
    Dominio =
        if parametros = null then {}
        else List.Transform(
            parametros,
            each Text.Upper(Text.From(_))
        ),

    Valor =
        if valor = null then null
        else Text.Upper(Text.From(valor)),

    Valido =
        if valor = null then true
        else List.Contains(Dominio, Valor),

    Ocorrencias =
        if Valido then null
        else
            {
                fxSchemaOcorrencia(
                    "DOMAIN",
                    contexto,
                    valor,
                    valor,
                    "Valor fora do domínio permitido.",
                    [Permitidos = parametros]
                )
            }

in
    [Valor = valor, Ocorrencias = Ocorrencias];
shared fxValidacaoSize = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let
    TamanhoEsperado =
        if parametros = null or List.IsEmpty(parametros) then null
        else Number.From(parametros{0}),

    TamanhoAtual =
        if valor = null then null
        else Text.Length(Text.From(valor)),

    Valido =
        valor = null
        or TamanhoEsperado = null
        or TamanhoAtual = TamanhoEsperado,

    Ocorrencias =
        if Valido then null
        else
            {
                fxSchemaOcorrencia(
                    "SIZE",
                    contexto,
                    valor,
                    valor,
                    "Tamanho inválido.",
                    [Esperado = TamanhoEsperado, Atual = TamanhoAtual]
                )
            }

in
    [Valor = valor, Ocorrencias = Ocorrencias];
shared fxValidacaoMin = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let
    Minimo =
        if parametros = null or List.IsEmpty(parametros) or parametros{0} = null then null
        else try Number.From(parametros{0}) otherwise null,

    Valor =
        if valor = null then null
        else try Number.From(valor) otherwise null,

    Valido =
        if valor = null then true
        else Minimo <> null and Valor <> null and Valor >= Minimo,

    Ocorrencias =
        if Valido then null
        else
            {
                fxSchemaOcorrencia(
                    "MIN",
                    contexto,
                    valor,
                    valor,
                    "Valor inferior ao mínimo permitido.",
                    [Minimo = Minimo, Valor = Valor]
                )
            }

in
    [Valor = valor, Ocorrencias = Ocorrencias];
shared fxValidacaoMax = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let
    Maximo =
        if parametros = null or List.IsEmpty(parametros) or parametros{0} = null then null
        else try Number.From(parametros{0}) otherwise null,

    Valor =
        if valor = null then null
        else try Number.From(valor) otherwise null,

    Valido =
        if valor = null then true
        else if Maximo = null then true
        else if Valor = null then false
        else Valor <= Maximo,

    Ocorrencias =
        if Valido then null
        else
            {
                fxSchemaOcorrencia(
                    "MAX",
                    contexto,
                    valor,
                    valor,
                    "Valor superior ao máximo permitido.",
                    [Maximo = Maximo, Valor = Valor]
                )
            }

in
    [Valor = valor, Ocorrencias = Ocorrencias];
shared fxValidacaoInterval = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let
    Minimo =
        if parametros = null or List.Count(parametros) < 1 or parametros{0} = null 
        then null 
        else Number.From(parametros{0}),

    Maximo =
        if parametros = null or List.Count(parametros) < 2 or parametros{1} = null 
        then null 
        else Number.From(parametros{1}),

    Valor =
        if valor = null then null 
        else Number.From(valor),

    Valido =
        if Valor = null then true
        else
            (Minimo = null or Valor >= Minimo)
            and
            (Maximo = null or Valor <= Maximo),

    Ocorrencias =
        if Valido then null
        else
            {
                fxSchemaOcorrencia(
                    "INTERVAL",
                    contexto,
                    valor,
                    valor,
                    "Valor fora do intervalo permitido.",
                    [Minimo = Minimo, Maximo = Maximo, Valor = Valor]
                )
            }

in
    [Valor = valor, Ocorrencias = Ocorrencias];
shared fxValidacaoEmail = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let
    Valido =
        if valor = null then true
        else
            let
                Texto = Text.From(valor),
                Partes = Text.Split(Texto, "@")
            in
                List.Count(Partes) = 2
                and Text.Length(Partes{0}) > 0
                and Text.Contains(Partes{1}, ".")
                and Text.PositionOf(Partes{1}, ".") > 0
                and Text.PositionOf(Partes{1}, ".") < Text.Length(Partes{1}) - 1,

    Ocorrencias =
        if Valido then null
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
    [Valor = valor, Ocorrencias = Ocorrencias];
shared fxValidacaoURL = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let
    URL =
        if valor = null then null
        else Text.Trim(Text.From(valor)),

    Protocolo =
        if URL = null then null
        else if Text.StartsWith(URL, "http://") then "http://"
        else if Text.StartsWith(URL, "https://") then "https://"
        else null,

    Dominio =
        if Protocolo = null then ""
        else Text.AfterDelimiter(URL, Protocolo),

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
        if Valido then null
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
    [Valor = valor, Ocorrencias = Ocorrencias];
shared fxValidacaoCEP = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let
    Ocorrencias =
        if valor = null then null
        else
            let
                CEP = Text.Trim(Text.From(valor)),
                Digitos = Text.Select(CEP, {"0".."9"}),
                FormatoSemMascara =
                    Text.Length(CEP) = 8
                    and Text.Length(Digitos) = 8,
                FormatoComMascara =
                    Text.Length(CEP) = 9
                    and Text.Length(Digitos) = 8
                    and Text.Middle(CEP, 5, 1) = "-",
                Valido = FormatoSemMascara or FormatoComMascara
            in
                if Valido then null
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
    [Valor = valor, Ocorrencias = Ocorrencias];
shared fxValidacaoCPF = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let
    CPF = if valor = null then null else Text.From(valor),
    PodeValidar = CPF <> null and Text.Length(CPF) = 11,

    Digitos =
        if PodeValidar then
            List.Transform(Text.ToList(CPF), each Number.FromText(_))
        else
            {},

    NaoRepetido =
        PodeValidar and List.Count(List.Distinct(Digitos)) > 1,

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
        if NaoRepetido then Number.Mod(SomaDV1 * 10, 11) else null,

    DV1 = if RestoDV1 = 10 then 0 else RestoDV1,

    SomaDV2 =
        if NaoRepetido then
            List.Sum(
                List.Transform(
                    {0..9},
                    each
                        (if _ = 9 then DV1 else Digitos{_})
                        *
                        (11 - _)
                )
            )
        else
            null,

    RestoDV2 =
        if NaoRepetido then Number.Mod(SomaDV2 * 10, 11) else null,

    DV2 = if RestoDV2 = 10 then 0 else RestoDV2,

    Valido =
        if valor = null then true
        else if not PodeValidar then false
        else if not NaoRepetido then false
        else
            DV1 = Digitos{9}
            and
            DV2 = Digitos{10},

    Ocorrencias =
        if Valido then null
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
    [Valor = valor, Ocorrencias = Ocorrencias];
shared fxValidacaoCNPJ = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let
    ValorNulo = valor = null,

    CNPJ =
        if ValorNulo then null
        else Text.Select(Text.From(valor), {"0".."9"}),

    Digitos =
        if ValorNulo then {}
        else List.Transform(Text.ToList(CNPJ), each Number.FromText(_)),

    TamanhoValido = List.Count(Digitos) = 14,
    NaoRepetido =
        TamanhoValido and List.Count(List.Distinct(Digitos)) > 1,

    PesosDV1 = {5,4,3,2,9,8,7,6,5,4,3,2},

    SomaDV1 =
        if NaoRepetido then
            List.Sum(
                List.Transform({0..11}, each Digitos{_} * PesosDV1{_})
            )
        else
            null,

    RestoDV1 =
        if NaoRepetido then Number.Mod(SomaDV1, 11) else null,

    DV1 =
        if NaoRepetido then
            if RestoDV1 < 2 then 0 else 11 - RestoDV1
        else
            null,

    PesosDV2 = {6,5,4,3,2,9,8,7,6,5,4,3,2},

    SomaDV2 =
        if NaoRepetido then
            List.Sum(
                List.Transform(
                    {0..12},
                    each
                        (if _ = 12 then DV1 else Digitos{_})
                        *
                        PesosDV2{_}
                )
            )
        else
            null,

    RestoDV2 =
        if NaoRepetido then Number.Mod(SomaDV2, 11) else null,

    DV2 =
        if NaoRepetido then
            if RestoDV2 < 2 then 0 else 11 - RestoDV2
        else
            null,

    Valido =
        if ValorNulo then true
        else
            NaoRepetido
            and
            DV1 = Digitos{12}
            and
            DV2 = Digitos{13},

    Ocorrencias =
        if Valido then null
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
    [Valor = valor, Ocorrencias = Ocorrencias];

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
        if valor = null then true
        else if valor is text then Text.Trim(valor) = ""
        else false,

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
        Valor = valor,
        Ocorrencias = Ocorrencias
    ];
shared fxNrmAplicar = /*
Entrada: Dados válidos (qaClientes → filtro OK)
   ↓
Processamento:
  1. Filtragem apenas de registros válidos
  2. Deduplicação por chave de negócio
  3. Resolução de relacionamentos
  4. Aplicação de regras complexas
  5. Enriquecimento com dados externos
   ↓
Saída: nrmClientes (dados normalizados)

Funções:
  - fxNrmAplicar() → Normalização completa
*/

(
    tabela as table,
    optional Schema as nullable text
) as table =>

let
    // RESPONSABILIDADE: Normalização de dados
    // (após STG + TRN + QA)
    
    // 1. Remover duplicatas por domínio
    //RegistrosUnicos = 
    //    Table.Distinct(Tabela),

    // 2. Resolver relacionamentos (se aplicável)
    ComRelacionamentos = 
        tabela,

    // 3. Enriquecimento com dados externos (customização por schema)
    Enriquecida = 
        ComRelacionamentos,

    // 4. Aplicar regras complexas de negócio
    Resultado = 
        Enriquecida

in
    Resultado;

//==============================================================================
// FUNÇÕES UTILITÁRIAS - Utility Functions
//==============================================================================

shared parTabelaCategoriasConsultasPQ = "tbSobreCategoriasConsultasPQ" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

//==============================================================================
// STAGING - STAGE (ETL Layer 1)
// Responsabilidade: Apenas estruturação e normalização básica
//==============================================================================
shared fxStgPreparar = 
(
    tabela as table,
    optional ignorarColunas as nullable list
)
as table =>

let
    // Configuração
    ColunasIgnoradas = ignorarColunas ?? {},
    ColunasTabela = Table.ColumnNames(tabela),
    ColunasEfetivas = List.Difference(ColunasTabela, ColunasIgnoradas),

    // Validação
    _ErroSemColunas = 
        if List.IsEmpty(ColunasTabela) then
            error "A tabela não possui colunas."
        else
            null,

    // 1. Remover linhas completamente vazias
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

    // 2. Normalizar nomes de colunas (trim, sem transformações)
    ColunasNormalizadas = 
        List.ReplaceMatchingItems(
            ColunasTabela,
            List.Transform(
                ColunasEfetivas,
                each {_, Text.Trim(_)}
            )
        ),

    // Validações de nomes
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

    TemDuplicidade = 
        List.Count(ColunasNormalizadas) <>
        List.Count(List.Distinct(ColunasNormalizadas)),

    _ErroDuplicidade = 
        if TemDuplicidade then
            error "A tabela possui nomes de colunas duplicados."
        else
            null,

    // 3. Aplicar renomeação
    Resultado = 
        Table.RenameColumns(
            TabelaSemLinhasVazias,
            List.Zip({ColunasTabela, ColunasNormalizadas}),
            MissingField.Ignore
        )

in
    Resultado;
shared fxStgAplicar = /*
Entrada: Dados brutos (srcClientes)
   ↓
Processamento:
  1. Remoção de linhas completamente vazias
  2. Normalização de nomes de colunas (trim, validação)
  3. Aplicação de tipos básicos
  4. Reordenação de colunas
   ↓
Saída: stgClientes (estruturado, tipado, ordenado)

Funções:
  - fxStgPreparar() → Preparação
  - fxStgAplicar() → Aplicação completa
*/

(
    Tabela as table,
    optional Schema as nullable text,
    optional ignorarColunas as nullable list
)
as table =>

let
    // Preparação estrutural
    Preparada = fxStgPreparar(Tabela, ignorarColunas),

    Pipeline = 
        if Schema = null then null 
        else try Record.Field(cfgPipeline, Schema) otherwise null,

    TemPipeline = Pipeline <> null,

    Resultado = 
        if not TemPipeline then
            Preparada
        else
            let
                // Tipos do pipeline
                Tipos = Pipeline[Tipos],
                Ordem = Pipeline[Ordem],

                // Transformações apenas de tipos básicos
                TransformacoesBasicas = 
                    List.RemoveNulls(
                        List.Transform(
                            Tipos,
                            (Def) =>
                                let
                                    Coluna = Def[Coluna],
                                    TipoDestino = Def[Tipo]
                                in
                                    if TipoDestino = type any then
                                        null
                                    else
                                        {
                                            Coluna,
                                            (v) => 
                                                if v = null then null
                                                else try fxConversor(v, TipoDestino) 
                                                     otherwise null,
                                            TipoDestino
                                        }
                        )
                    ),

                // Aplicar tipos
                ComTipos = 
                    if List.IsEmpty(TransformacoesBasicas) then
                        Preparada
                    else
                        Table.TransformColumns(
                            Preparada,
                            TransformacoesBasicas,
                            null,
                            MissingField.Ignore
                        ),

                // Reordenar colunas
                Final = 
                    Table.ReorderColumns(
                        ComTipos,
                        Ordem,
                        MissingField.Ignore
                    )
            in
                Final
in
    Resultado;

//==============================================================================
// TRATAMENTOS - TRANSFORM (ETL Layer 2)
// Responsabilidade: Tratamentos e limpeza de dados
//==============================================================================
shared fxTratamentoNormalizeBasic = (valor as any, optional parametros as nullable any) as any =>

let
    Resultado =
        if valor = null then null
        else
            let
                Texto = Text.From(valor),
                Limpo = Text.Clean(Texto),
                Aparado = Text.Trim(Limpo),
                Normalizado =
                    Text.Combine(
                        List.Select(
                            Text.SplitAny(Aparado, " "),
                            each _ <> ""
                        ),
                        " "
                    )
            in
                if Normalizado = "" then null else Normalizado

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

shared fxObjetoIdentificarPeloNome = (
    objeto as text,
    source as text
)
as record =>

let

//--------------------------------------------------------------------------
// Prefixos cadastrados
//--------------------------------------------------------------------------

    Prefixos =
        cfgPrefixosPowerQuery,

//--------------------------------------------------------------------------
// Nome normalizado
//--------------------------------------------------------------------------

    Objeto =
        Text.Lower(objeto),

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

                    each Text.StartsWith(
                        Objeto,
                        _
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

//==============================================================================
// VALIDAÇÕES - Validation Functions
//==============================================================================


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
                        fxPipelineCompilar
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

                List.RemoveNulls(

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

                                else if not Record.FieldOrDefault(Definicao, "Ativo", true) then
                                    null

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

shared fxPipelineCompilar = 
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

shared fxTrnAplicar = /*
Entrada: Dados estruturados (stgClientes)
   ↓
Processamento:
  1. Compilação de operadores do schema
  2. Aplicação de tratamentos (TRIM, UPPER, etc.)
  3. Conversão refinada de tipos
  4. Execução otimizada (única passagem)
   ↓
Saída: trnClientes (dados tratados)

Operadores Disponíveis (20+):
  Texto: TRIM, UPPER, LOWER, PROPER, CLEAN, SINGLESPACE, DIGITS, ALPHANUMERIC
  Numérico: ABS, ROUND
  Normalização: NORMALIZEBASIC, NORMALIZETYPE
  Docs: CPF, CNPJ, CEP

Funções:
  - fxTrnAplicar() → Aplicação de transformações
  - fxTrnCompilarTratamentosPorColuna() → Compilação otimizada
*/

(
    Tabela as table,
    optional Schema as nullable text
)
as table =>

let

    // Obter pipeline compilado
    Pipeline =
        if
            Schema = null
            or Text.Trim(Schema) = ""
        then
            [
                TratamentosPorColuna = [],
                TiposPorColuna = []
            ]
        else
            try
                Record.Field(
                    cfgPipeline,
                    Schema
                )
            otherwise
                error Error.Record(
                    "Schema não encontrado",
                    Text.Format(
                        "O schema '#{0}' não está configurado.",
                        {Schema}
                    )
                ),

    TratamentosPorColuna =
        Pipeline[TratamentosPorColuna],

    TiposPorColuna =
        Pipeline[TiposPorColuna],

    // Compilar transformações
    Transformacoes =
        List.RemoveNulls(
            List.Transform(
                Record.FieldNames(TratamentosPorColuna),
                (Coluna) =>
                    let
                        Operadores =
                            Record.Field(
                                TratamentosPorColuna,
                                Coluna
                            ),

                        TipoPipeline =
                            Record.FieldOrDefault(
                                TiposPorColuna,
                                Coluna,
                                type any
                            )

                    in
                        if List.IsEmpty(Operadores) then
                            null
                        else
                            let
                                TratadorCompilado =
                                    fxTrnCompilarTratamentosPorColuna(
                                        Operadores
                                    ),

                                TransformadorFinal =
                                    (valor) =>
                                        let
                                            ResultadoTratamento =
                                                TratadorCompilado(
                                                    valor
                                                )
                                        in
                                            if ResultadoTratamento = null then
                                                null
                                            else if Value.Is(ResultadoTratamento, TipoPipeline) then
                                                ResultadoTratamento
                                            else
                                                try
                                                    fxConversor(
                                                        ResultadoTratamento,
                                                        TipoPipeline
                                                    )
                                                otherwise
                                                    ResultadoTratamento
                            in
                                if TipoPipeline = type any then
                                    {
                                        Coluna,
                                        TransformadorFinal
                                    }
                                else
                                    {
                                        Coluna,
                                        TransformadorFinal,
                                        TipoPipeline
                                    }
            )
        ),

    // Aplicar transformações
    Resultado =
        if List.IsEmpty(Transformacoes) then
            Tabela
        else
            Table.TransformColumns(
                Tabela,
                Transformacoes,
                null,
                MissingField.Ignore
            )

in
    Resultado;

shared fxTrnCompilarTratamentosPorColuna = (
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

//==============================================================================
// VALIDAÇÕES - QUALITY ASSURANCE (ETL Layer 3)
// Responsabilidade: Validações estruturais, semânticas e de negócio
//==============================================================================



shared fxQaValidar = /*
Entrada: Dados transformados (trnClientes)
   ↓
Processamento:
  1. Compilação de validadores do schema
  2. Execução de validações (não deleta, apenas marca)
  3. Acúmulo de ocorrências de erro
  4. Adição de status (_QA_Status: OK, AVISO, ERRO)
   ↓
Saída: qaClientes (com _QA_Status e _QA_Ocorrencias)

Validadores Disponíveis (15+):
  Obrigatoriedade: REQUIRED
  Documentos: CPFVAL, CNPJVAL, CEPVAL
  Internet: EMAIL, URL, DOMAIN
  Intervalo: MIN, MAX, INTERVAL
  Listas: LIST, SIZE

Funções:
  - fxQaValidar() → Validação completa
  - fxQaFiltrarPorStatus() → Filtro por status
  - fxQaExtrairProblemas() → Extração para auditoria
  - fxQaCompilarValidacoesPorColuna() → Compilação otimizada
*/

(
    Tabela as table,
    optional Schema as nullable text
)
as table =>

let
    // Obter pipeline com early exit
    Pipeline =
        if Schema = null or Text.Trim(Schema) = "" then
            [ ValidacoesPorColuna = [], TiposPorColuna = [] ]
        else
            try Record.Field(cfgPipeline, Schema)
            otherwise 
                error Error.Record(
                    "Schema não encontrado",
                    Text.Format("O schema '#{0}' não está configurado.", {Schema})
                ),

    ValidacoesPorColuna = Pipeline[ValidacoesPorColuna],
    TiposPorColuna = Pipeline[TiposPorColuna],
    ColunasValidacao = Record.FieldNames(ValidacoesPorColuna),

    // Early exit se não houver validações
    Resultado =
        if List.IsEmpty(ColunasValidacao) then
            Table.AddColumn(Tabela, "_QA_Status", each "OK", type text)
        else
            let
                // Pré-compilação otimizada dos validadores (uma única vez)
                ValidadoresPorColuna =
                    Record.FromList(
                        List.Transform(
                            ColunasValidacao,
                            (Coluna) => 
                                fxQaCompilarValidacoesPorColuna(
                                    Record.Field(ValidacoesPorColuna, Coluna),
                                    Record.FieldOrDefault(TiposPorColuna, Coluna, type any),
                                    Coluna
                                )
                        ),
                        ColunasValidacao
                    ),

                // Pré-compilação do cache de severidades (uma única vez)
                CacheSeveridades = 
                    let
                        Nomes = Record.FieldNames(cfgParametrosSeveridades)
                    in
                        Record.FromList(
                            List.Transform(
                                Nomes, 
                                each Record.FieldOrDefault(
                                    Record.Field(cfgParametrosSeveridades, _), 
                                    "Bloqueia", 
                                    false
                                )
                            ),
                            Nomes
                        ),

                // Única passagem pela tabela com lógica otimizada
                TabelaComQA = Table.AddColumn(
                    Tabela,
                    "_QA",
                    each 
                        let
                            Linha = _,
                            
                            // Acumulação otimizada de ocorrências usando Accumulate + early append
                            Ocorrencias = List.Accumulate(
                                ColunasValidacao,
                                {},
                                (estado, Coluna) => 
                                    let
                                        Validador = Record.Field(ValidadoresPorColuna, Coluna),
                                        Valor = Record.FieldOrDefault(Linha, Coluna, null),
                                        ResultadoVal = Validador(Valor)
                                    in
                                        if ResultadoVal = null or List.IsEmpty(ResultadoVal) then 
                                            estado 
                                        else 
                                            estado & ResultadoVal
                            ),
                            
                            // Cálculo de status com cache
                            Status = 
                                if List.IsEmpty(Ocorrencias) then 
                                    "OK"
                                else 
                                    let
                                        TemErro = List.MatchesAny(
                                            Ocorrencias,
                                            (o) => Record.FieldOrDefault(CacheSeveridades, o[Severidade], false)
                                        )
                                    in
                                        if TemErro then "ERRO" else "AVISO"
                        in
                            [ Status = Status, Ocorrencias = if List.IsEmpty(Ocorrencias) then null else Ocorrencias ],
                    type record
                ),

                // Expansão final (mais eficiente que duas AddColumn separadas)
                ResultadoFinal = Table.ExpandRecordColumn(
                    TabelaComQA,
                    "_QA",
                    {"Status", "Ocorrencias"},
                    {"_QA_Status", "_QA_Ocorrencias"}
                )
            in
                ResultadoFinal

in
    Resultado;

// Função auxiliar para filtrar por status QA


shared fxQaFiltrarPorStatus = (
    Tabela as table,
    optional Status as nullable text
) as table =>

let
    FiltroStatus = Status ?? "OK",
    
    // Filtrar por status
    Filtrada = 
        Table.SelectRows(
            Tabela,
            each [_QA_Status] = FiltroStatus
        ),

    // Remover colunas de controle QA
    Resultado = 
        Table.RemoveColumns(
            Filtrada,
            {"_QA_Status", "_QA_Ocorrencias"},
            MissingField.Ignore
        )
in
    Resultado;

// Função para extrair registros com problemas


shared fxQaExtrairProblemas = (
    Tabela as table
) as table =>

let
    // Filtrar apenas registros com problemas
    ComProblemas = 
        Table.SelectRows(
            Tabela,
            each [_QA_Status] <> "OK"
        )
in
    ComProblemas;

shared fxQaCompilarValidacoesPorColuna = (
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

//==============================================================================
// NORMALIZAÇÃO - NORMALIZE (ETL Layer 4)
// Responsabilidade: Estrutura de dados, deduplicação, enriquecimento
//==============================================================================

shared trnClientes = let
    Fonte = stgClientes,
    Resultado = fxTrnAplicar(Fonte, parTabelaClientes)
in
    Resultado;

shared qaClientes = let
    Fonte = trnClientes,
    qa = fxQaValidar(Fonte, parTabelaClientes)
in
    qa;

shared trnProdutos = let
    Fonte = stgProdutos,
    Resultado = fxTrnAplicar(Fonte, parTabelaProdutos)
in
    Resultado;

shared qaProdutos = let
    Fonte = trnProdutos,
    qa = fxQaValidar(Fonte, parTabelaProdutos)

    // Usar apenas dados válidos
    // Validos = fxQaFiltrarPorStatus(qa, "OK"),

    // Extrair problemas para auditoria
    // Problemas = fxQaExtrairProblemas(qa)

in
    qa;

shared trnVendas = let
    Fonte = stgVendas,
    Resultado = fxTrnAplicar(Fonte, parTabelaVendas)
in
    Resultado;

shared qaVendas = let
    Fonte = trnVendas,
    qa = fxQaValidar(Fonte, parTabelaVendas)
in
    qa;

shared tstTesteSTG = // Antes
let Resultado = fxStgAplicar(srcClientes, "tbClientes") in Resultado

// Esperado: Dados estruturados, tipos básicos, sem vazios
;

shared tstTesteTRN = let 
    STG = tstTesteSTG,
    Resultado = fxTrnAplicar(STG, "tbClientes")
in 
    Resultado

// Esperado: Dados tratados (TRIM, UPPER, etc. aplicados)
;

shared tstTesteQA = // Novo
let 
    STG = tstTesteSTG,
    TRN = tstTesteTRN,
    Resultado = fxQaValidar(TRN, "tbClientes")
in 
    Resultado

// Esperado: Colunas _QA_Status e _QA_Ocorrencias adicionadas
;

shared tstTesteNRM = let 
    STG = tstTesteSTG,
    TRN = tstTesteTRN,
    QA = tstTesteQA,
    Valido = fxQaFiltrarPorStatus(QA, "OK"),
    Resultado = fxNrmAplicar(Valido, "tbClientes")
in 
    Resultado

// Esperado: Dados deduplicated, enriquecidos, sem problemas
;

shared tstTesteAntesDepois = // Comparar registros antes/depois
let
    Antes = tstTesteSTG,
    Depois = tstTesteNRM,
    Diferenca = Table.NestedJoin(Antes, {"CPF"}, Depois, {"CPF"}, "x", JoinKind.FullOuter)
in
    Diferenca;

shared tstTesteProblemasEncontrados = // Extrair problemas encontrados
let
    ProblemasDetectados = fxQaExtrairProblemas(qaClientes),
    _QA_OcorrenciasExpandido = Table.ExpandListColumn(ProblemasDetectados, "_QA_Ocorrencias"),
    _QA_OcorrenciasComMensagem = Table.ExpandRecordColumn(_QA_OcorrenciasExpandido, "_QA_Ocorrencias", {"Coluna", "Mensagem"}, {"Coluna", "Mensagem"})
in
    _QA_OcorrenciasComMensagem;

shared tstClientes1M = let
    Schema = "tstClientes1M",
    Fonte = stgDados,
    Stage = fxStgAplicar(Fonte, Schema),
    Transform = fxTrnAplicar(Stage, Schema),
    Qa = fxQaValidar(Transform, Schema),
    Valida = fxQaFiltrarPorStatus(Qa, "OK"),
    Normalize = fxNrmAplicar(Valida, Schema)

in
    Normalize;

shared fxConversaoParaTexto = (Valor as any, Contexto as record) as record =>
    try
        let
            ValorTexto = Text.From(Valor),
            Normalizado = fxNormalizarTextoBasico(ValorTexto)
        in
            [
                Status = "OK",
                Valor = Normalizado,
                Tipo = "TEXT",
                Contexto = Contexto
            ]
    otherwise
        [
            Status = "ERRO",
            Valor = Valor,
            Tipo = "TEXT",
            Mensagem = "Falha ao converter para texto",
            Contexto = Contexto
        ];

shared fxConversaoParaNumero = (Valor as any, Contexto as record) as record =>
    try
        let
            // Se for texto, limpar caracteres especiais
            ValorLimpo = if Value.Type(Valor) = "text" then
                Text.Remove(Valor, {" ", ",", "$", "R$", "%"})
            else
                Valor,
            
            // Tentar conversão
            ValorNumerico = Number.From(ValorLimpo)
        in
            [
                Status = "OK",
                Valor = ValorNumerico,
                Tipo = "NUMBER",
                Contexto = Contexto
            ]
    otherwise
        [
            Status = "ERRO",
            Valor = Valor,
            Tipo = "NUMBER",
            Mensagem = "Falha ao converter para número",
            Contexto = Contexto
        ];

shared fxConversaoParaBooleano = (Valor as any, Contexto as record) as record =>
    try
        let
            // Normalizar entrada
            ValorNormalizado = if Value.Type(Valor) = "text" then
                Text.Upper(Text.Trim(Valor))
            else if Value.Type(Valor) = "logical" then
                Valor
            else
                Text.Upper(Text.From(Valor)),
            
            // Verificar valores verdadeiros
            ValoresVerdadeiros = {"TRUE", "SIM", "YES", "1", "V", "VERDADEIRO", "ATIVO", "ON"},
            ValoresFalsos = {"FALSE", "NAO", "NÃO", "NO", "0", "F", "FALSO", "INATIVO", "OFF"},
            
            ValorBooleano = if List.Contains(ValoresVerdadeiros, ValorNormalizado) then
                true
            else if List.Contains(ValoresFalsos, ValorNormalizado) then
                false
            else
                error "Valor booleano inválido"
        in
            [
                Status = "OK",
                Valor = ValorBooleano,
                Tipo = "LOGICAL",
                Contexto = Contexto
            ]
    otherwise
        [
            Status = "ERRO",
            Valor = Valor,
            Tipo = "LOGICAL",
            Mensagem = "Falha ao converter para booleano",
            Contexto = Contexto
        ];

shared fxConversaoParaData = (Valor as any, Contexto as record) as record =>
    try
        let
            ValorData = Date.From(Valor)
        in
            [
                Status = "OK",
                Valor = ValorData,
                Tipo = "DATE",
                Contexto = Contexto
            ]
    otherwise
        [
            Status = "ERRO",
            Valor = Valor,
            Tipo = "DATE",
            Mensagem = "Falha ao converter para data",
            Contexto = Contexto
        ]
;

shared fxConversaoParaHora = (Valor as any, Contexto as record) as record =>
    try
        let
            ValorHora = Time.From(Valor)
        in
            [
                Status = "OK",
                Valor = ValorHora,
                Tipo = "TIME",
                Contexto = Contexto
            ]
    otherwise
        [
            Status = "ERRO",
            Valor = Valor,
            Tipo = "TIME",
            Mensagem = "Falha ao converter para hora",
            Contexto = Contexto
        ];

shared fxConversaoParaDataHora = (Valor as any, Contexto as record) as record =>
    try
        let
            ValorDataHora = DateTime.From(Valor)
        in
            [
                Status = "OK",
                Valor = ValorDataHora,
                Tipo = "DATETIME",
                Contexto = Contexto
            ]
    otherwise
        [
            Status = "ERRO",
            Valor = Valor,
            Tipo = "DATETIME",
            Mensagem = "Falha ao converter para data/hora",
            Contexto = Contexto
        ]
;

shared fxConversaoParaMoeda = (Valor as any, Contexto as record) as record =>
    try
        let
            // Limpar e converter
            ValorLimpo = if Value.Type(Valor) = "text" then
                Text.Remove(Valor, {" ", "R$", "%"})
            else
                Valor,
            
            ValorNumerico = Number.From(ValorLimpo)
        in
            [
                Status = "OK",
                Valor = ValorNumerico,
                Tipo = "CURRENCY",
                Formato = "R$",
                Contexto = Contexto
            ]
    otherwise
        [
            Status = "ERRO",
            Valor = Valor,
            Tipo = "CURRENCY",
            Mensagem = "Falha ao converter para moeda",
            Contexto = Contexto
        ]
;

shared fxConversaoParaPercentual = (Valor as any, Contexto as record) as record =>
    try
        let
            // Limpar e converter
            ValorLimpo = if Value.Type(Valor) = "text" then
                Text.Remove(Valor, {" ", "%"})
            else
                Valor,
            
            ValorNumerico = Number.From(ValorLimpo),
            // Se valor > 1, assume percentual (ex: 50 = 50%)
            ValorPercentual = if ValorNumerico > 1 then
                ValorNumerico / 100
            else
                ValorNumerico
        in
            [
                Status = "OK",
                Valor = ValorPercentual,
                Tipo = "PERCENTAGE",
                PercentualExibicao = ValorPercentual * 100,
                Contexto = Contexto
            ]
    otherwise
        [
            Status = "ERRO",
            Valor = Valor,
            Tipo = "PERCENTAGE",
            Mensagem = "Falha ao converter para percentual",
            Contexto = Contexto
        ]
;

shared fxNormalizarTextoBasico = (
    Texto as text,
    optional Opcoes as nullable record
) as text =>
    let
        // Opções padrão
        OpcoesPadrao = Opcoes ?? [
            Trim = true,
            Clean = true,
            SingleSpace = true,
            RemoveNullBytes = false,
            RemoveLineBreaks = true
        ],
        
        // 1. Remover null bytes se configurado
        Passo1 = if OpcoesPadrao[RemoveNullBytes] then
            Text.Remove(Texto, {Character.FromNumber(0)})
        else
            Texto,
        
        // 2. Remover quebras de linha se configurado
        Passo2 = if OpcoesPadrao[RemoveLineBreaks] then
            Text.Replace(
                Text.Replace(Passo1, "#(lf)", " "),
                "#(cr)",
                " "
            )
        else
            Passo1,
        
        // 3. Trim (remover espaços início/fim)
        Passo3 = if OpcoesPadrao[Trim] then
            Text.Trim(Passo2)
        else
            Passo2,
        
        // 4. Clean (remover caracteres de controle)
        Passo4 = if OpcoesPadrao[Clean] then
            Text.Clean(Passo3)
        else
            Passo3,
        
        // 5. Single space (remover espaços múltiplos)
        Passo5 = if OpcoesPadrao[SingleSpace] then
            Text.Combine(
                List.Select(
                    Text.SplitAny(Passo4, " "),
                    each _ <> ""
                ),
                " "
            )
        else
            Passo4
    in
        Passo5
;

shared fxTratamentoReplace = (valor as any, optional parametros as nullable any) as any =>
    //{"antigo","novo"}
    if valor = null then null else
        Text.Replace(
            Text.From(valor),
            Text.From(parametros{0}),
            Text.From(parametros{1})
        );

shared fxTratamentoLeft = (valor as any, optional parametros as nullable any) as any =>
    // quantidade
    if valor = null then
        null
    else
        Text.Start(
            Text.From(valor),
            Number.From(parametros{0})
        );

shared fxTratamentoRight = (valor as any, optional parametros as nullable any) as any =>
    // quantidade
    if valor = null then null else
        Text.End(Text.From(valor), Number.From(parametros{0}));

shared fxTratamentoMid = (valor as any, optional parametros as nullable any) as any =>
    // {posição, quantidade}
    if valor = null then null else
        Text.Middle(
            Text.From(valor),
            Number.From(parametros{0}),
            Number.From(parametros{1})
        );

shared fxTratamentoBefore = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else
        Text.BeforeDelimiter(Text.From(valor), Text.From(parametros{0}));

shared fxTratamentoAfter = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else
        Text.AfterDelimiter(Text.From(valor), Text.From(parametros{0}));

shared fxTratamentoPrefix = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else
        Text.From(parametros{0}) & Text.From(valor);

shared fxTratamentoSuffix = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else
        Text.From(valor) & Text.From(parametros{0});

shared fxTratamentoPadLeft = (valor as any, optional parametros as nullable any) as any =>
    // {tamanho, caractere}
    if valor = null then null else
        Text.PadStart(
            Text.From(valor),
            Number.From(parametros{0}),
            Text.From(parametros{1})
        );

shared fxTratamentoPadRight = (valor as any, optional parametros as nullable any) as any =>
    // {tamanho, caractere}
    if valor = null then null else
        Text.PadEnd(
            Text.From(valor),
            Number.From(parametros{0}),
            Text.From(parametros{1})
        );

shared fxTratamentoRemoveChars = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else
        Text.Remove(Text.From(valor), parametros);

shared fxTratamentoKeepChars = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else
        Text.Select(Text.From(valor), parametros);

shared fxTratamentoRemoveAccents = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then
        null
    else
        let
            Texto = Text.From(valor),

            Mapa = {
                {"Á","A"},{"À","A"},{"Â","A"},{"Ã","A"},{"Ä","A"},
                {"á","a"},{"à","a"},{"â","a"},{"ã","a"},{"ä","a"},
                {"É","E"},{"È","E"},{"Ê","E"},{"Ë","E"},
                {"é","e"},{"è","e"},{"ê","e"},{"ë","e"},
                {"Í","I"},{"Ì","I"},{"Î","I"},{"Ï","I"},
                {"í","i"},{"ì","i"},{"î","i"},{"ï","i"},
                {"Ó","O"},{"Ò","O"},{"Ô","O"},{"Õ","O"},{"Ö","O"},
                {"ó","o"},{"ò","o"},{"ô","o"},{"õ","o"},{"ö","o"},
                {"Ú","U"},{"Ù","U"},{"Û","U"},{"Ü","U"},
                {"ú","u"},{"ù","u"},{"û","u"},{"ü","u"},
                {"Ç","C"},{"ç","c"},
                {"Ñ","N"},{"ñ","n"}
            },

            Resultado =
                List.Accumulate(
                    Mapa,
                    Texto,
                    (Estado, Item) =>
                        Text.Replace(
                            Estado,
                            Item{0},
                            Item{1}
                        )
                )
        in
            Resultado;

shared fxTratamentoPunctuation = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else
        Text.Remove(Text.From(valor), ".,;:!?()[]{}<>/\|-_""'");