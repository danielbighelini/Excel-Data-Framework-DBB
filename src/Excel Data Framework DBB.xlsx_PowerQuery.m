// Power Query from: Excel Data Framework DBB.xlsx
// Pathname: c:\Users\daniel-bighelini\OneDrive\Documentos\04_Projetos\Excel-Data-Framework-DBB\src\Excel Data Framework DBB.xlsx
// Extracted: 2026-08-20T14:42:35.078Z

section Section1;

shared srcObjetosPowerQuery = // Lista os nomes de todas as queries na seção atual do arquivo M.
Record.FieldNames(
    #sections[Section1]
)
;

shared srcCategoriasPowerQuery = let
    Fonte = srcWorkbook{[Name=parTabelaCategoriasConsultasPQ]}[Content]
in
    Fonte;

shared srcWorkbook = // Workbook Excel atual bufferizado; ponto de entrada de todas as leituras de tabelas.
Table.Buffer(Excel.CurrentWorkbook())
;

shared srcParametrosExcel = 
// Lê tbParametros; retorna tabela vazia se a tabela não existir.
let
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

shared srcOperadores = // Catálogo de todos os operadores (tratamentos + validações) com ponteiros para as funções de implementação.
let
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
            Parâmetros = nullable record
        ],
        {
            // Tratamentos básicos
            {"TRIM", "Remove espaços em branco do início e do final do texto.", fxTratamentoTrim, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"UPPER", "Converte todo o texto para letras maiúsculas.", fxTratamentoUpper, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"LOWER", "Converte todo o texto para letras minúsculas.", fxTratamentoLower, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"PROPER", "Converte a primeira letra de cada palavra para maiúscula.", fxTratamentoProper, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"CLEAN", "Remove caracteres não imprimíveis do texto.", fxTratamentoClean, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"EMPTYTONULL", "Converte valores vazios ou em branco para null.", fxTratamentoEmptyToNull, type any, type any, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"NULLTOEMPTY", "Converte valores nulos (null) em textos vazios.", fxTratamentoNullToEmpty, type any, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"SINGLESPACE", "Substitui múltiplos espaços consecutivos por apenas um espaço.", fxTratamentoSingleSpace, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"DIGITS", "Mantém apenas os dígitos numéricos do texto.", fxTratamentoDigits, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"ALPHANUMERIC", "Mantém apenas letras e números, removendo símbolos e pontuações.", fxTratamentoAlphaNumeric, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"ABS", "Retorna o valor absoluto (positivo) de um número.", fxTratamentoAbs, type number, type number, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"ROUND", "Arredonda um número decimal para a quantidade de casas especificadas.", fxTratamentoRound, type number, type number, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"NORMALIZEBASIC", "Remove caracteres de controle, elimina espaços excedentes no início, fim e entre palavras, retornando um texto normalizado ou nulo quando vazio.", fxTratamentoNormalizeBasic, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"NUMBER", "Converte textos contendo valores numéricos em um número, reconhecendo automaticamente sinais, moeda e separadores decimal e de milhar.", fxTratamentoNumber, type text, type number, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"CPF", "Formata ou extrai apenas os números para o padrão de CPF.", fxTratamentoDigits, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"CNPJ", "Formata ou extrai apenas os números para o padrão de CNPJ.", fxTratamentoDigits, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"CEP", "Formata ou extrai apenas os números para o padrão de CEP.", fxTratamentoDigits, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"PHONE", "Formata ou extrai apenas os números para o padrão de telefone.", fxTratamentoDigits, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            
            // Tratamentos avançados
            {"REPLACE", "Substitui todas as ocorrências de um texto por outro.", fxTratamentoReplace, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"LEFT", "Mantém os N primeiros caracteres do texto.", fxTratamentoLeft, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"RIGHT", "Mantém os N últimos caracteres do texto.", fxTratamentoRight, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"MID", "Extrai uma quantidade de caracteres a partir de uma posição específica.", fxTratamentoMid, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"BEFORE", "Extrai o texto localizado antes de um delimitador informado.", fxTratamentoBefore, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"AFTER", "Extrai o texto localizado após um delimitador informado.", fxTratamentoAfter, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"ADDPREFIX", "Adiciona um prefixo ao início do texto.", fxTratamentoAddPrefix, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"ADDSUFFIX", "Adiciona um sufixo ao final do texto.", fxTratamentoAddSuffix, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"REMOVEPREFIX", "Remove o primeiro prefixo encontrado no início do texto.", fxTratamentoRemovePrefix, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"REMOVESUFFIX", "Remove o primeiro sufixo encontrado no final do texto.", fxTratamentoRemoveSuffix, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"PADLEFT", "Completa o texto à esquerda até atingir o comprimento especificado.", fxTratamentoPadLeft, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"PADRIGHT", "Completa o texto à direita até atingir o comprimento especificado.", fxTratamentoPadRight, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"REMOVECHARS", "Remove todos os caracteres pertencentes a uma lista informada.", fxTratamentoRemoveChars, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"KEEPCHARS", "Mantém apenas os caracteres pertencentes a uma lista informada.", fxTratamentoKeepChars, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"REMOVEACCENTS", "Remove acentos e caracteres diacríticos do texto.", fxTratamentoRemoveAccents, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"REMOVEPUNCTUATION", "Remove todos os sinais de pontuação do texto.", fxTratamentoRemovePunctuation, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},
            {"KEEPTEXT", "Mantém apenas caracteres e espaços do texto.", fxTratamentoKeepText, type text, type text, false, true, CategoriaTratamento, SeveridadeAviso, null},

            // Validações
            {"REQUIRED", "Valida se um campo obrigatório foi preenchido.", fxValidacaoRequired, type any, type any, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"CPFVAL", "Valida se o número de CPF informado é matematicamente válido.", fxValidacaoCPF, type text, type text, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"CNPJVAL", "Valida se o número de CNPJ informado é matematicamente válido.", fxValidacaoCNPJ, type text, type text, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"CEPVAL", "Valida se o formato do CEP informado está correto.", fxValidacaoCEP, type text, type text, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"PHONEVAL", "Valida se o valor possui um formato válido de telefone brasileiro (fixo ou celular), aceitando números com ou sem máscara e com ou sem código do país (+55).", fxValidacaoPhone, type text, type text, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"EMAIL", "Valida se a estrutura do endereço de e-mail está correta.", fxValidacaoEmail, type text, type text, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"URL", "Valida se a estrutura do endereço web (URL) está correta.", fxValidacaoURL, type text, type text, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"LIST", "Valida se o valor pertence a uma lista de opções permitidas.", fxValidacaoList, type any, type any, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"SIZE", "Valida se o tamanho ou comprimento do dado está dentro do limite.", fxValidacaoSize, type text, type text, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"MIN", "Valida se o valor é maior ou igual ao limite mínimo permitido.", fxValidacaoMin, type number, type number, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"MAX", "Valida se o valor é menor ou igual ao limite máximo permitido.", fxValidacaoMax, type number, type number, false, true, CategoriaValidacao, SeveridadeErro, null},
            {"INTERVAL", "Valida se o valor está dentro de um intervalo numérico ou temporal específico.", fxValidacaoInterval, type number, type number, false, true, CategoriaValidacao, SeveridadeErro, null}
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

shared srcParametrosFeriados = let
    Fonte = srcWorkbook{[Name=parTabelaParametrosFeriados]}[Content]
in
    Fonte;
shared srcParametrosBooleanos = // Mapa string → logical; aceita representações em pt-BR e en-US.
let
    Fonte = srcWorkbook{[Name=parTabelaParametrosBooleanos]}[Content]
in
    Fonte;

shared srcParametrosTipos = //==============================================================================
// TIPOS DE DADOS - Configuração Central
//==============================================================================
// Mapa nome → tipo M; aceita nomes em pt-BR e inglês para compatibilidade.
let
    Fonte = srcWorkbook{[Name=parTabelaParametrosTipos]}[Content]
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
shared parTabelaParametros = "tbParametros" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared cfgCultura = 
// Record {Separadores, Símbolos, Formatações} derivados da cultura ativa; usado por fxConversor e fxTratamentoNumber.
let
    Parametro =
        (Nome as text, Padrao as any) as any =>
            Record.FieldOrDefault(
                cfgParametros,
                Nome,
                [Valor = Padrao]
            )[Valor],

    Cultura = parCulturaBootstrap,

    TextoNumero =
        try
            Number.ToText(
                1234.5,
                "#,##0.0",
                Cultura
            )
        otherwise
            error "Cultura inválida.",

    SeparadorMilhar =
        Text.At(TextoNumero, 1),

    SeparadorDecimal =
        Text.At(
            TextoNumero,
            Text.Length(TextoNumero) - 2
        ),

    SimboloMoeda =
        Parametro(
            "Cultura_Simbolo_Moeda",
            "R$"
        ),

    SimboloPercentual =
        Parametro(
            "Cultura_Simbolo_Percentual",
            "%"
        ),

    NegativoParenteses =
        fxParseBooleano(
            Parametro(
                "Cultura_Negativo_Parenteses",
                true
            ),
            Cultura
        ),

    CaracteresPermitidosNumero =
        List.Buffer(
            List.Combine(
                {
                    {"0".."9"},
                    {SeparadorDecimal, SeparadorMilhar},
                    {"-","+","(",")"}
                }
            )
        ),

    SeparadorLista =
        Parametro(
            "Cultura_Separador_Lista",
            ";"
        ),

    SeparadorParametros =
        Parametro(
            "Cultura_Separador_Parametros",
            ","
        ),

    FormatoData = 
        Parametro(
            "Cultura_Formato_Data",
            "dd/mm/aaaa"
        ),

    FormatoHora = 
        Parametro(
            "Cultura_Formato_Hora",
            "hh:mm:ss"
        ),

    FormatoDataHora = 
        Parametro(
            "Cultura_Formato_DataHora",
            "dd/mm/aaaa hh:mm"
        ),

    PrimeiroDiaSemana =
        Parametro(
            "Cultura_Primeiro_Dia_Semana",
            "Domingo"
        ),

    Resultado =
        [
            Cultura = Cultura,
            SeparadorDecimal = SeparadorDecimal,
            SeparadorMilhar = SeparadorMilhar,
            SimboloMoeda = SimboloMoeda,
            SimboloPercentual = SimboloPercentual,
            NegativoParenteses = NegativoParenteses,
            CaracteresPermitidosNumero = CaracteresPermitidosNumero,
            SeparadorLista = SeparadorLista,
            SeparadorParametros = SeparadorParametros,
            FormatoData = FormatoData,
            FormatoHora = FormatoHora,
            FormatoDataHora = FormatoDataHora,
            PrimeiroDiaSemana = PrimeiroDiaSemana
        ]
in
    Resultado
;

shared cfgVersaoInfo = // Versão do framework; incremente a cada alteração que quebre compatibilidade.
let
    Parametro =
        (Nome as text, Padrao as any) as any =>
            Record.FieldOrDefault(
                cfgParametros,
                Nome,
                [Valor = Padrao]
            )[Valor],

    VersaoInfo =
        [
            VersaoFramework = Parametro("Framework_Versao", "1.0.0"),
            DataAtualizacaoVersao = Parametro("Framework_Atualizacao", null)
        ]
in
    VersaoInfo;

shared cfgParametros =
// Materializa a tabela consolidada em um Record para permitir
// acesso O(1) aos parâmetros através de Record.Field.
let
    cfgParametros = 
        Record.FromTable(
            Table.Buffer(stgParametros)
        )
in
    cfgParametros;

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
    optional valorPadrao as nullable any,
    optional culture as nullable text,
    optional config as nullable record
)
as any =>

let

    Parametros =
        cfgParametros,

    NomeParametro =
        Text.Upper(parametro),

    ExisteOverride =
        config <> null
        and Record.HasFields(
            config,
            NomeParametro
        ),

    Parametro =
        Record.FieldOrDefault(
            Parametros,
            NomeParametro
        ),

    ExisteParametro =
        Parametro <> null,

    Tipo =
        if ExisteParametro then
            Record.FieldOrDefault(
                Parametro,
                "Tipo",
                type any
            )
        else
            type any,

    Obrigatorio =
        if ExisteParametro then
            Record.FieldOrDefault(
                Parametro,
                "Obrigatório",
                true
            )
        else
            false,

    Permitidos =
        if ExisteParametro then
            Record.FieldOrDefault(
                Parametro,
                "Permitidos"
            )
        else
            null,

    Valor =
        if ExisteOverride then
            Record.Field(
                config,
                NomeParametro
            )
        else if ExisteParametro then
            let
                v =
                    Record.FieldOrDefault(
                        Parametro,
                        "Valor"
                    )
            in
                if v = null then
                    valorPadrao
                else
                    v
        else
            valorPadrao,

    Resultado =
        if
            not ExisteParametro
            and not ExisteOverride
            and valorPadrao = null
        then

            error Error.Record(
                "Parâmetro inexistente",
                Text.Format(
                    "O parâmetro '#{0}' não foi encontrado.",
                    {
                        parametro
                    }
                ),
                [
                    Parametro = parametro,
                    Disponiveis =
                        Text.Combine(
                            List.Sort(
                                Record.FieldNames(
                                    Parametros
                                )
                            ),
                            ", "
                        )
                ]
            )

        else if
            ExisteParametro
            and Obrigatorio
            and Valor = null
        then

            error Error.Record(
                "Parâmetro obrigatório",
                Text.Format(
                    "O parâmetro '#{0}' é obrigatório e não possui valor.",
                    {
                        parametro
                    }
                ),
                [
                    Parametro = parametro
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
                            culture ?? "pt-BR"
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

                else if
                    Permitidos <> null
                    and not List.Contains(
                        Permitidos,
                        Conversao[Value]
                    )
                then

                    error Error.Record(
                        "Valor inválido",
                        Text.Format(
                            "O valor '#{0}' não é permitido para o parâmetro '#{1}'.",
                            {
                                Conversao[Value],
                                parametro
                            }
                        ),
                        [
                            Parametro = parametro,
                            Valor = Conversao[Value],
                            Permitidos = Permitidos
                        ]
                    )

                else
                    Conversao[Value]

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

shared stgObjetosExcel = let

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
                Categoria = text
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
                Categoria = nullable text
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

shared stgTabelasExcel = let
    Fonte =
        stgObjetosExcel,

    LinhasFiltradas =
        Table.SelectRows(
            Fonte,
            each not Text.Contains([Nome], "!") and not Text.StartsWith([Nome], "diag")
        ),
    
    ComSchema =
        Table.AddColumn(
            LinhasFiltradas,
            "Schema",
            each Table.Schema(srcWorkbook{[Name = [Nome]]}[Content])
        ),

    ComColunas =
        Table.AddColumn(
            ComSchema,
            "Columns",
            each [Schema][Name],
            type list
        ),

    Resultado =
        Table.AddColumn(
            ComColunas,
            "ColumnCount",
            each Table.RowCount([Schema]),
            Int64.Type
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
            and not List.Contains(parObjetosPowerQueryIgnorados, [Nome])

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
                        // O codigo abaixo foi desabilitado pelo fato de gerar consultas ciclicas
                        // e nao estar sendo necessario no framework
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
                            type logical,
                            parCulturaBootstrap
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
                                [Tipo],
                                parCulturaBootstrap
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

    NormalizarName =
        Table.TransformColumns(
            Fonte,
            {
                {
                    "Parâmetro",
                    Text.Upper,
                    type text
                }
            }
        ),

    Distintos =
        Table.Distinct(
            NormalizarName,
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
    Resultado;

shared fxParametroIdentificarTipo = (tipo as nullable text) as nullable type =>
// Converte nome de tipo em texto (ex: "TEXTO", "DATA") no type M correspondente via cfgTiposDados.
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
    else if Record.HasFields(cfgTipos, Nome) then
        Record.Field(cfgTipos, Nome)
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
                            Record.FieldNames(cfgTipos)
                        ),
                        ", "
                    )
            ]
        );


shared cfgTipos = let

    Fonte =
        Table.Buffer(stgParametrosTipos),

    Tipos =
        [
            LOGICAL = type logical,
            DATE = type date,
            DATETIME = type datetime,
            DATETIMEZONE = type datetimezone,
            DURATION = type duration,
            TIME = type time,
            LIST = type list,
            NUMBER = type number,
            INT64 = Int64.Type,
            ANY = type any,
            TEXT = type text
        ],

    cfgTipos =
        Record.Combine(
            List.Transform(
                Table.ToRecords(Fonte),
                each
                    Record.FromList(
                        {
                            Record.Field(Tipos, [Código]),
                            Record.Field(Tipos, [Código])
                        },
                        {
                            [Descrição],
                            [Código]
                        }
                    )
            )
        )

in

    cfgTipos;

shared cfgTiposObjetos = srcTiposObjetos;

shared cfgTiposObjetosPorNome = let
    Vals = Record.FieldValues(cfgTiposObjetos)
in
    Record.FromList(Vals, List.Transform(Vals, each [Nome]));
// Mapa Kind → metadados de tipo M (Nome, Type, Categoria, IsStructured).
shared srcTiposObjetos = [
    Table = [
        Kind = "Table",
        Nome = "Tabela",
        Type = type table,
        Categoria = "Estruturado"
    ],

    Record = [
        Kind = "Record",
        Nome = "Registro",
        Type = type record,
        Categoria = "Estruturado"
    ],

    List = [
        Kind = "List",
        Nome = "Lista",
        Type = type list,
        Categoria = "Estruturado"
    ],

    Function = [
        Kind = "Function",
        Nome = "Função",
        Type = type function,
        Categoria = "Executável"
    ],

    Action = [
        Kind = "Action",
        Nome = "Ação",
        Type = type action,
        Categoria = "Executável"
    ],

    Text = [
        Kind = "Text",
        Nome = "Texto",
        Type = type text,
        Categoria = "Escalar"
    ],

    Logical = [
        Kind = "Logical",
        Nome = "Lógico",
        Type = type logical,
        Categoria = "Escalar"
    ],

    Int64 = [
        Kind = "Int64",
        Nome = "Inteiro 64 bits",
        Type = Int64.Type,
        Categoria = "Numérico"
    ],

    Number = [
        Kind = "Number",
        Nome = "Número",
        Type = type number,
        Categoria = "Numérico"
    ],

    Date = [
        Kind = "Date",
        Nome = "Data",
        Type = type date,
        Categoria = "Data e Hora"
    ],

    Time = [
        Kind = "Time",
        Nome = "Hora",
        Type = type time,
        Categoria = "Data e Hora"
    ],

    DateTime = [
        Kind = "DateTime",
        Nome = "Data e Hora",
        Type = type datetime,
        Categoria = "Data e Hora"
    ],

    DateTimeZone = [
        Kind = "DateTimeZone",
        Nome = "Data e Hora com Fuso",
        Type = type datetimezone,
        Categoria = "Data e Hora"
    ],

    Duration = [
        Kind = "Duration",
        Nome = "Duração",
        Type = type duration,
        Categoria = "Data e Hora"
    ],

    Binary = [
        Kind = "Binary",
        Nome = "Binário",
        Type = type binary,
        Categoria = "Binário"
    ],

    Type = [
        Kind = "Type",
        Nome = "Tipo",
        Type = type type,
        Categoria = "Especial"
    ],

    Null = [
        Kind = "Null",
        Nome = "Nulo",
        Type = type null,
        Categoria = "Especial"
    ],

    Any = [
        Kind = "Any",
        Nome = "Qualquer",
        Type = type any,
        Categoria = "Especial"
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

shared stgOperadores = 
// Mescla srcOperadores com overrides de tbParametrosTratamentos e tbParametrosValidacoes; resolve Ativo/Padrão/Severidade efetivos.
let
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
                {"Coluna", type text},
                {"Código", type text},
                {"Ordem", Int64.Type},
                {"Ativo", type logical}
            }
        ),

    Tratamento =
        Table.TransformColumns(
            Tipos,
            {
                {"Coluna", each Text.Trim(_), type text},
                {"Código", each Text.Upper(Text.Trim(_)), type text}
            }
        ),

    Filtro =
        Table.SelectRows(
            Tratamento,
            each
                [Ativo] and
                [Código] <> null and
                [Código] <> "" and
                [Coluna] <> null and
                [Coluna] <> ""
        ),

    RemoverDuplicados =
        Table.Distinct(
            Filtro,
            {"Código"}
        ),

    Ordenar =
        Table.Sort(
            RemoverDuplicados,
            {
                {"Ordem", Order.Ascending},
                {"Coluna", Order.Ascending}
            }
        )

in
    Ordenar;

shared stgParametrosFeriados = let
    Fonte = srcParametrosFeriados,

    Tipos =
        Table.TransformColumnTypes(
            Fonte,
            {
                {"Nome", type text},
                {"Código", type text},
                {"Regularidade", type text},
                {"Referência", type text},
                {"Offset", Int64.Type},
                {"Dia", Int64.Type},
                {"Mês", Int64.Type},
                {"Abrangência", type text},
                {"PontoFacultativo", type logical},
                {"Expediente", type text},
                {"Estado", type text},
                {"Município", type text},
                {"VigênciaInicial", type date},
                {"VigênciaFinal", type date},
                {"Ativo", type logical}
            }
        ),

    Tratamento =
        Table.TransformColumns(
            Tipos,
            {
                {"Nome", each Text.Trim(_), type text},
                {"Código", each Text.Upper(Text.Trim(_)), type text},
                {"Regularidade", each Text.Upper(Text.Trim(_)), type text},
                {"Referência", each if _ = null then null else Text.Upper(Text.Trim(_)), type text},
                {"Abrangência", each Text.Trim(_), type text},
                {"Expediente", each Text.Trim(_), type text},
                {"Estado", each if _ = null then null else Text.Upper(Text.Trim(_)), type text},
                {"Município", each if _ = null then null else Text.Trim(_), type text}
            }
        ),

    FiltrarValidos =
        Table.SelectRows(
            Tratamento,
            each
                [Ativo] and
                [Código] <> null and
                [Código] <> ""
        ),

    RemoverDuplicados =
        Table.Distinct(
            FiltrarValidos,
            {"Código"}
        )
    
in
    RemoverDuplicados;

shared stgParametrosBooleanos = let

    Fonte =
        srcParametrosBooleanos,

    LinhasValidas =
        Table.SelectRows(
            Fonte,
            each
                [Descrição] <> null
                and Text.Trim(Text.From([Descrição])) <> ""
                and [Código] <> null
        ),

    Booleanos =
        Table.TransformColumns(
            LinhasValidas,
            {
                {
                    "Descrição",
                    each Text.Upper(Text.Trim(Text.From(_)))
                }
            }
        ),

    BooleanosTipados =
        Table.TransformColumnTypes(
            Booleanos,
            {
                {"Código", type logical},
                {"Descrição", type text}
            }
        ),

    Distintos =
        Table.Distinct(
            BooleanosTipados,
            {"Código", "Descrição"}
        )

in

    Distintos;

shared stgParametrosTipos = let

    Fonte =
        srcParametrosTipos,

    LinhasValidas =
        Table.SelectRows(
            Fonte,
            each
                [Descrição] <> null
                and Text.Trim(Text.From([Descrição])) <> ""
                and [Código] <> null
                and Text.Trim(Text.From([Código])) <> ""
        ),

    Normalizados =
        Table.TransformColumns(
            LinhasValidas,
            {
                {
                    "Descrição",
                    each Text.Upper(Text.Trim(Text.From(_))),
                    type text
                },
                {
                    "Código",
                    each Text.Upper(Text.Trim(Text.From(_))),
                    type text
                }
            }
        ),

    Distintos =
        Table.Distinct(
            Normalizados,
            {"Código", "Descrição"}
        )

in

    Distintos;

shared stgSchema = 
// Schema ativo normalizado e filtrado (Ativo=true); base para cfgSchema e cfgPipeline.

let

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

    // Garante compatibilidade com versões de tbSchema sem a coluna Chave.
    ComChave =
        if Table.HasColumns(LinhasValidas, "Chave") then LinhasValidas
        else Table.AddColumn(LinhasValidas, "Chave", each false, type logical),

    Normalizado =

        Table.TransformColumns(

            ComChave,

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

                },

                {
                    "Chave",
                    each fxConversor(_, type logical),
                    type nullable logical
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
                "Ativo",
                "Chave"

            }

        )

in

    Resultado
;

shared cfgTiposBooleanos = let
    cfgTiposBooleanos =
        Record.FromTable(
            Table.RenameColumns(
                Table.SelectColumns(
                    Table.Buffer(stgParametrosBooleanos),
                    {"Descrição", "Código"}
                ),
                {
                    {"Descrição", "Name"},
                    {"Código", "Value"}
                }
            )
        )
in
    cfgTiposBooleanos;


shared cfgConversores = srcConversores;
shared srcConversores = [
    Any = (v as any, culture as text) as any => v,
    Text = (v as any, culture as text) as any => Text.From(v, culture),
    List = (v as any, culture as text) as any => fxListaNormalizar(v),
    Int64 = (v as any, culture as text) as any => Int64.From(v, culture),
    Number = (v as any, culture as text) as any => Number.From(v, culture),
    Date = (v as any, culture as text) as any => Date.From(v, culture),
    DateTime = (v as any, culture as text) as any => DateTime.From(v, culture),
    DateTimeZone = (v as any, culture as text) as any => DateTimeZone.From(v, culture),
    Time = (v as any, culture as text) as any => Time.From(v, culture),
    Duration = (v as any, culture as text) as any => Duration.From(v),
    Logical = (v as any, culture as text) as any => fxParseBooleano(v, culture),
    Unknown = (v as any, culture as text) as any => v
];
shared fxListaNormalizar = (
    valor as any,
    optional separador as nullable text,
    optional removerVazios as nullable logical,
    optional trim as nullable logical
)
as nullable list =>

let
    // Converte texto delimitado ou lista em lista normalizada; remove vazios e duplicatas por padrão.
    Separador = separador ?? ";",
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
shared fxParseBooleano = // Converte representações textuais (SIM/NÃO/TRUE/FALSE/1/0) em logical via cfgTiposBooleanos.
(valor as any, optional culture as nullable text) as nullable logical =>
let
    Culture = culture ?? parCulturaBootstrap,
    Nome =
        if valor = null then null
        else Text.Upper(Text.Trim(Text.From(valor, Culture))),
    Sentinela = "__##NAO_ENCONTRADO##__",
    Achado = if Nome = null then null
             else Record.FieldOrDefault(cfgTiposBooleanos, Nome, Sentinela)
in
    if Nome = null then null
    else if Achado <> Sentinela then Achado
    else error Error.Record(
        "Valor lógico inválido",
        Text.Format("O valor '#{0}' não é uma representação válida de lógico.", {valor}),
        [Valor = valor,
         ValoresPermitidos = Text.Combine(List.Sort(Record.FieldNames(cfgTiposBooleanos)), ", ")]
    )
;

shared fxConversor = (valor as any, tipo as nullable type, optional culture as nullable text) as any =>
// Hub central de conversão de tipo M; seleciona o conversor pelo tipo e aplica com a cultura informada.
let
    Culture = if culture = null then parCulturaBootstrap else culture,
    TipoDestino = if tipo = null then type any else tipo,
    Chave = fxTipoParaTexto(TipoDestino),
    ChaveFinal  =
        if Chave = "Unknown" then
            error Error.Record(
                "Tipo não suportado",
                "A função não possui conversor para o tipo solicitado.",
                [Tipo = TipoDestino]
            )
        else
            Chave,

    Conversor = Record.Field(cfgConversores, ChaveFinal),
    Resultado = Conversor(valor, Culture)
in
    Resultado;

shared fxTipoParaTexto = (tipo as type) as text =>
let
    Resultado =
        if tipo = type any then "Any"
        else if tipo = type text then "Text"
        else if tipo = type list then "List"
        else if tipo = Int64.Type then "Int64"
        else if tipo = type number then "Number"
        else if tipo = type date then "Date"
        else if tipo = type datetime then "Datetime"
        else if tipo = type datetimezone then "Datetimezone"
        else if tipo = type time then "Time"
        else if tipo = type duration then "Duration"
        else if tipo = type logical then "Logical"
        else "Unknown"
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

    //-------------------------------------------------------------------------
    // Fonte: apenas os nomes das queries
    //-------------------------------------------------------------------------
    Fonte = srcObjetosPowerQuery,

    //-------------------------------------------------------------------------
    // Transformação mínima: Nome + Categoria
    //-------------------------------------------------------------------------
    ComCategoria = Table.AddColumn(
        Table.FromList(Fonte, Splitter.SplitByNothing(), {"Nome"}),
        "Categoria",
        each 
            let
                Prefixo = fxObjetoIdentificarPrefixo([Nome]),
                CategoriaRecord = Record.FieldOrDefault(
                    cfgCategoriasPowerQuery,
                    Prefixo,
                    [Categoria = "Qualquer"]
                )
            in
                CategoriaRecord[Categoria],
        type text
    ),

    //-------------------------------------------------------------------------
    // Ordenação final
    //-------------------------------------------------------------------------
    Resultado = Table.Sort(
        ComCategoria,
        {{"Categoria", Order.Ascending}, {"Nome", Order.Ascending}}
    )

in
    Resultado;

shared diagTabelasExcel = let

    //-------------------------------------------------------------------------
    // Fonte: objetos Excel filtrados
    //-------------------------------------------------------------------------
    Fonte = stgObjetosExcel,

    LinhasFiltradas = Table.SelectRows(
        Fonte,
        each not Text.Contains([Nome], "!") and not Text.StartsWith([Nome], "diag")
    ),

    //-------------------------------------------------------------------------
    // Extração otimizada: apenas nomes de colunas
    //-------------------------------------------------------------------------
    ComColunas = Table.AddColumn(
        LinhasFiltradas,
        "Columns",
        each 
            let
                Conteudo = srcWorkbook{[Name = [Nome]]}[Content],
                ColunasNomes = Table.ColumnNames(Conteudo)
            in
                ColunasNomes,
        type list
    ),

    //-------------------------------------------------------------------------
    // Transformar lista em texto separado por ponto e vírgula
    //-------------------------------------------------------------------------
    ColunasTexto = Table.TransformColumns(
        ComColunas,
        {
            {
                "Columns",
                each Text.Combine(_, ";"),
                type text
            }
        }
    ),

    //-------------------------------------------------------------------------
    // Selecionar apenas Nome e Colunas
    //-------------------------------------------------------------------------
    Resultado = Table.SelectColumns(
        ColunasTexto,
        {"Nome", "Columns"}
    ),

    //-------------------------------------------------------------------------
    // Renomear para consistência com a versão original
    //-------------------------------------------------------------------------
    ColunasRenomeadas = Table.RenameColumns(
        Resultado,
        {
            {"Nome", "Tabela"},
            {"Columns", "Colunas"}
        }
    )

in
    ColunasRenomeadas;

shared stgClientes = let
    Fonte = srcClientes,
    Preparada = fxStgAplicar(Fonte, parTabelaClientes)
in
    Preparada;

shared stgProdutos = let
    Fonte = srcProdutos,
    Preparada = fxStgAplicar(Fonte, parTabelaProdutos),
    Resultado = Preparada
in
    Resultado;

shared stgVendas = let
    Fonte = srcVendas,
    Preparada = fxStgAplicar(Fonte, parTabelaVendas),
    Resultado = Preparada
in
    Resultado;

shared nrmClientes = let
    Fonte = qaClientes,
    Validada = fxQaFiltrarPorStatus(Fonte, "OK"),
    Normalizada = fxNrmAplicar(Validada, parTabelaClientes)
in
    Normalizada;

shared dimClientes = let
    Fonte =
        nrmClientes,

    Chaves =
        Table.AddIndexColumn(
            Fonte,
            "IDCliente",
            1,
            1,
            Int64.Type
        ),

    Reordenada =
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

    Buffer =
        Table.Buffer(
            Reordenada
        )
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
    
    // Resolve a chave da dimensão Calendário.
    Calendario =
        Table.NestedJoin(
            ProdutoExpandido,
            {"Data"},
            dimCalendario,
            {"Data"},
            "_Cal",
            JoinKind.LeftOuter
        ),
    CalendarioExpandido =
        Table.ExpandTableColumn(
            Calendario,
            "_Cal",
            {"IDData"}
        ),
    
    // Seleciona apenas as colunas do modelo dimensional.
    Colunas =
        Table.SelectColumns(
            CalendarioExpandido,
            {
                "IDData",
                "IDCliente",
                "IDProduto",
                "Quantidade",
                "ValorUnitário",
                "ValorTotal"
            },
            MissingField.Ignore
        ),
    
    // Organiza as colunas da tabela fato.
    ColunasReordenadas =
        Table.ReorderColumns(
            Colunas,
            {
                "IDData",
                "IDCliente",
                "IDProduto",
                "Quantidade",
                "ValorUnitário",
                "ValorTotal"
            },
            MissingField.Ignore
        )
in
    ColunasReordenadas;

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
    optional Url as nullable text,
    optional Config as nullable record
)
as any =>

let
    Endpoint =
        if Url = null then

            fxParametro(
                "REST_Endpoint_Base",
                null,
                null,
                Config
            )
            &
            fxParametro(
                "REST_Endpoint_Path",
                "",
                null,
                Config
            )

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
    optional Headers as nullable record,
    optional Config as nullable record
)
as binary =>

let
    BaseUrl =
        fxParametro(
            "REST_Endpoint_Base",
            null,
            null,
            Config
        ),

    Path =
        fxParametro(
            "REST_Endpoint_Path",
            "",
            null,
            Config
        ),

    Timeout =
        fxParametro(
            "REST_Timeout",
            30,
            null,
            Config
        ),

    Tentativas =
        fxParametro(
            "REST_Tentativas",
            3,
            null,
            Config
        ),

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
                    #"Content-Type" =
                        "text/xml; charset=utf-8"
                ],
                cfgRESTHeaders,
                if Headers = null then
                    []
                else
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
                            Content =
                                Text.ToBinary(
                                    Corpo,
                                    TextEncoding.Utf8
                                ),
                            Timeout =
                                #duration(
                                    0,
                                    0,
                                    0,
                                    Timeout
                                ),
                            IsRetry =
                                Tentativa > 0,
                            ManualStatusHandling =
                                StatusReprocessar
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
                        Value.Metadata(
                            Binario
                        )[Response.Status]
                    otherwise
                        200,

            DeveRepetir =
                (
                    Resposta[HasError]
                    or List.Contains(
                        StatusReprocessar,
                        Status
                    )
                )
                and Tentativa < Tentativas,

            Resultado =
                if DeveRepetir then

                    Function.InvokeAfter(
                        () =>
                            @Requisitar(
                                Tentativa + 1
                            ),
                        #duration(
                            0,
                            0,
                            0,
                            Number.Power(
                                2,
                                Tentativa
                            )
                        )
                    )

                else if Resposta[HasError] then

                    error Resposta[Error]

                else if
                    List.Contains(
                        StatusReprocessar,
                        Status
                    )
                then

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
    optional Headers as nullable record,
    optional Config as nullable record
)
as binary =>

let
    BaseUrl =
        fxParametro(
            "REST_Endpoint_Base",
            null,
            null,
            Config
        ),

    Path =
        if RelativePath = null then
            fxParametro(
                "REST_Endpoint_Path",
                null,
                null,
                Config
            )
        else
            RelativePath,

    Timeout =
        fxParametro(
            "REST_Timeout",
            30,
            null,
            Config
        ),

    Tentativas =
        fxParametro(
            "REST_Tentativas",
            3,
            null,
            Config
        ),

    Cabecalhos =
        if Headers = null then
            cfgRESTHeaders
        else
            Record.Combine(
                {
                    cfgRESTHeaders,
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
                            Timeout =
                                #duration(
                                    0,
                                    0,
                                    0,
                                    Timeout
                                ),
                            IsRetry =
                                Tentativa > 0,
                            ManualStatusHandling =
                                StatusReprocessar
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
                        Value.Metadata(
                            Binario
                        )[Response.Status]
                    otherwise
                        200,

            DeveRepetir =
                (
                    Resposta[HasError]
                    or List.Contains(
                        StatusReprocessar,
                        Status
                    )
                )
                and Tentativa < Tentativas,

            Resultado =
                if DeveRepetir then

                    Function.InvokeAfter(
                        () =>
                            @Requisitar(
                                Tentativa + 1
                            ),
                        #duration(
                            0,
                            0,
                            0,
                            Number.Power(
                                2,
                                Tentativa
                            )
                        )
                    )

                else if Resposta[HasError] then

                    error Resposta[Error]

                else if
                    List.Contains(
                        StatusReprocessar,
                        Status
                    )
                then

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

shared fxOrigemRESTConteudo = (
    optional Config as nullable record
)
as any =>

let
    Protocolo =
        Text.Upper(
            fxParametro(
                "REST_Protocolo",
                null,
                null,
                Config
            )
        ),

    Resultado =
        if Protocolo = "REST" then

            fxRESTRequest(
                null,
                null,
                Config
            )

        else if Protocolo = "ODATA" then

            fxODataRequest(
                null,
                Config
            )

        else if Protocolo = "SOAP" then

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

shared fxOrigem = (
    optional Config as nullable record
)
as any =>

let
    CarregarDadosExternos =
        fxParametro(
            "Dados_Carregar",
            true,
            null,
            Config
        ),

    FonteDados =
        Text.Upper(
            fxParametro(
                "Dados_Fonte",
                null,
                null,
                Config
            )
        ),

    UsarFontePrincipal =
        Config = null,

    Resultado =
        if not CarregarDadosExternos then

            #table({}, {})

        else if FonteDados = "ARQUIVOS" then

            if UsarFontePrincipal then
                srcArquivos
            else
                fxOrigemArquivos(Config)

        else if FonteDados = "REST" then

            if UsarFontePrincipal then
                srcREST
            else
                fxOrigemREST(Config)

        else if FonteDados = "SGBD" then

            if UsarFontePrincipal then
                srcSGBD
            else
                fxOrigemSGBD(Config)

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

shared fxOrigemArquivos = (
    optional Config as nullable record
)
as record =>

let
    Origem =
        Text.Upper(
            fxParametro(
                "Arquivo_Origem",
                null,
                null,
                Config
            )
        ),
        
    Resultado =
        if Origem = "LOCAL" or Origem = "REMOTA" then

            let
                Caminho =
                    fxResolverCaminho(
                        fxParametro(
                            "Local_Pasta",
                            null,
                            null,
                            Config
                        )
                    ),

                Arquivos =
                    fxConectorLocal(
                        Caminho[Pasta]
                    )

            in
                [
                    Arquivos = Arquivos,
                    Arquivo = Caminho[Arquivo]
                ]

        else if Origem = "SHAREPOINT" then

            fxConectorSharePoint(
                fxParametro(
                    "Sharepoint_Site",
                    null,
                    null,
                    Config
                ),
                fxCaminhoSharePoint(
                    fxParametro(
                        "Sharepoint_Pasta",
                        null,
                        null,
                        Config
                    )
                )
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

shared fxOrigemSGBD = (
    optional Config as nullable record
)
as table =>

let
    TipoSGBD =
        Text.Upper(
            fxParametro(
                "SGBD_Tipo",
                null,
                null,
                Config
            )
        ),

    Host =
        fxParametro(
            "SGBD_Host",
            null,
            null,
            Config
        ),

    Banco =
        fxParametro(
            "SGBD_Banco",
            null,
            null,
            Config
        ),

    SQL =
        fxParametro(
            "SGBD_SQL",
            null,
            null,
            Config
        ),

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

shared fxOrigemREST = (
    optional Config as nullable record
)
as any =>

let
    Conteudo =
        fxOrigemRESTConteudo(
            Config
        ),

    Resultado =
        if Value.Is(
            Conteudo,
            Binary.Type
        ) then

            let
                Formato =
                    Text.Upper(
                        fxParametro(
                            "REST_Formato",
                            null,
                            null,
                            Config
                        )
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
shared fxOrigemComoTabela = (Valor as any) as table =>
// Normaliza qualquer valor (table, record, list, scalar) em tabela de uma coluna.
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

shared cfgRESTHeaders = srcRESTHeaders;

shared fxLeitorCSV = (
    Content as binary,
    optional pDelimiter as nullable text,
    optional pEncoding as nullable number,
    optional pQuoteStyle as nullable number,
    optional pPromoteHeaders as nullable logical
)
as table =>

let
    Fonte =
        Csv.Document(
            Content,
            [
                Delimiter =
                    pDelimiter
                    ?? ";",

                Encoding =
                    pEncoding
                    ?? 65001,

                QuoteStyle =
                    pQuoteStyle
                    ?? QuoteStyle.Csv
            ]
        ),

    PromoverCabecalhos =
        pPromoteHeaders
        ?? true,

    Cabecalhos =
        if
            not PromoverCabecalhos
            or Table.IsEmpty(Fonte)
        then
            Fonte
        else
            Table.PromoteHeaders(
                Fonte,
                [
                    PromoteAllScalars = true
                ]
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
                            and Text.Trim(
                                Text.From(Valor)
                            ) <> ""
                    )
                )
        )

in
    SemLinhasVazias
;

shared fxLeitorExcel = (
    Content as binary,
    optional pUseHeaders as nullable any,
    optional pDelayTypes as nullable logical
)
as table =>

let
    Resultado =
        Excel.Workbook(
            Content,
            pUseHeaders,
            pDelayTypes ?? true
        )

in
    Resultado;

shared fxLeitorJSON = (
    Content as binary,
    optional pEncoding as nullable number
)
as any =>

let
    Resultado =
        Json.Document(
            Content,
            pEncoding ?? TextEncoding.Utf8
        )

in
    Resultado;

shared fxLeitorXML = (
    Content as binary,
    optional pOptions as nullable record,
    optional pEncoding as nullable number
)
as table =>

let
    Resultado =
        Xml.Tables(
            Content,
            pOptions,
            pEncoding
        )

in
    Resultado;

shared fxLeitorPDF = (
    Content as binary,
    optional pOptions as nullable record
)
as table =>

let
    Resultado =
        Pdf.Tables(
            Content,
            pOptions
        )

in
    Resultado;

shared fxLeitorArquivo = (
    Content as binary,
    optional pFormatoArquivo as nullable text,
    optional pDelimitadorArquivo as nullable text,
    optional pCodificadorArquivo as nullable number,
    optional pQuoteStyle as nullable number,
    optional pPromoteHeaders as nullable logical
)
as any =>

let
    FormatoArquivo =
        Text.Upper(
            Text.From(
                pFormatoArquivo
                    ?? fxParametro("Arquivo_Formato")
            )
        ),

    DelimitadorArquivo =
        pDelimitadorArquivo
            ?? fxParametro(
                "Arquivo_Delimitador",
                ";"
            ),

    CodificadorArquivo =
        pCodificadorArquivo
            ?? fxParametro(
                "Arquivo_Codificador",
                65001
            ),

    Resultado =
        if FormatoArquivo = "CSV" then

            fxLeitorCSV(
                Content,
                DelimitadorArquivo,
                CodificadorArquivo,
                pQuoteStyle,
                pPromoteHeaders
            )

        else if FormatoArquivo = "EXCEL" then

            fxLeitorExcel(
                Content
            )

        else if FormatoArquivo = "JSON" then

            fxLeitorJSON(
                Content,
                CodificadorArquivo
            )

        else if FormatoArquivo = "XML" then

            fxLeitorXML(
                Content,
                CodificadorArquivo
            )

        else if FormatoArquivo = "PDF" then

            fxLeitorPDF(
                Content
            )

        else

            error Error.Record(
                "Formato de arquivo não suportado",
                Text.Format(
                    "Não existe leitor implementado para o formato '#{0}'.",
                    {
                        FormatoArquivo
                    }
                ),
                [
                    Formato = FormatoArquivo
                ]
            )

in
    Resultado;

shared srcArquivos = let
    FonteDados =
        Text.Upper(
            fxParametro(
                "Dados_Fonte"
            )
        ),

    Resultado =
        if FonteDados <> "ARQUIVOS" then

            #table({}, {})

        else

            let
                Origem =
                    fxOrigemArquivos(),

                FormatoArquivo =
                    Text.Upper(
                        Text.From(
                            fxParametro(
                                "Arquivo_Formato",
                                ""
                            )
                        )
                    ),

                DelimitadorArquivo =
                    fxParametro(
                        "Arquivo_Delimitador",
                        ";"
                    ),

                CodificadorArquivo =
                    fxParametro(
                        "Arquivo_Codificador",
                        65001
                    ),

                QuoteStyleArquivo =
                    fxParametro(
                        "Arquivo_QuoteStyle",
                        QuoteStyle.Csv
                    ),

                PromoverCabecalhos =
                    fxParametro(
                        "Arquivo_Promover_Cabecalhos",
                        true
                    ),

                ArquivosFiltrados =
                    fxFiltrarArquivos(
                        Origem[Arquivos],
                        FormatoArquivo,
                        Origem[Arquivo]
                    ),

                Dados =
                    Table.AddColumn(
                        ArquivosFiltrados,
                        "Dados",
                        each
                            fxLeitorArquivo(
                                [Content],
                                FormatoArquivo,
                                DelimitadorArquivo,
                                CodificadorArquivo,
                                QuoteStyleArquivo,
                                PromoverCabecalhos
                            ),
                        type any
                    )

            in
                Dados

in
    Resultado;

shared cfgFormatosArquivos = let
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

shared cfgSeveridades = let

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
            "Calendario_Data_Inicial",
            null
        ),

    DataFinalManual =
        fxParametro(
            "Calendario_Data_Final",
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

shared cfgFeriados = let
    Fonte = Table.Buffer(stgParametrosFeriados),

    Fixos =
        Table.SelectRows(
            Fonte,
            each [Regularidade] = "FIXO"
        ),

    Moveis =
        Table.SelectRows(
            Fonte,
            each [Regularidade] = "MÓVEL"
        ),

    PorCodigo =
        Record.FromList(
            Table.ToRecords(Fonte),
            Fonte[Código]
        ),

    Configuracao =
        [
            Todos = Fonte,
            Fixos = Fixos,
            Moveis = Moveis,
            PorCodigo = PorCodigo
        ]
in
    Configuracao;

shared cfgFeriadosIndices = let
    Anos =
        {Date.Year(cfgCalendario[DataInicial])..Date.Year(cfgCalendario[DataFinal])},

    Indices =
        List.Transform(
            Anos,
            each
                let
                    Ano = _,

                    Tabela =
                        fxFeriadoGerar(Ano),

                    Chaves =
                        List.Transform(
                            Tabela[Data],
                            each Date.ToText(_, "yyyy-MM-dd")
                        ),

                    Valores =
                        List.Transform(
                            Table.ToRecords(Tabela),
                            each [
                                Existe = true,
                                Nome = [Nome],
                                Código = [Código],
                                Abrangência = [Abrangência],
                                PontoFacultativo = [PontoFacultativo],
                                Expediente = [Expediente],
                                Estado = [Estado],
                                Município = [Município]
                            ]
                        )
                in
                    Record.FromList(
                        Valores,
                        Chaves
                    )
        ),

    Resultado =
        Record.FromList(
            Indices,
            List.Transform(
                Anos,
                each Text.From(_)
            )
        )
in
    Resultado;

shared cfgCalendarioIntervalos = // Estrutura de configuração utilizada para definir o intervalo de datas
// que será considerado na dimensão de calendário.
//
// Quando a tabela é informada como null, a função utiliza automaticamente
// o primeiro e o último dia do ano atual.
//
// Quando uma tabela é informada, a função analisa a coluna indicada e
// determina a menor e a maior data encontrada.
//
// Para adicionar um intervalo específico de uma tabela, inclua uma nova
// chamada à função no registro abaixo, informando a tabela e a coluna de data.

let
    // Lista de intervalos que serão utilizados pela dimensão de calendário.
    //
    // null = utiliza automaticamente o intervalo do ano atual.
    // "Data" = coluna utilizada para determinar o intervalo quando uma
    // tabela específica for informada.
    cfgIntervalos = {
        fxCalendarioIntervaloData(null, "Data")
    }
in
    cfgIntervalos;

shared cfgCalendarioAtributos = // Lido uma vez; capturado nas closures de todos os atributos fiscais.
let
    InicioFiscal = fxParametro("Calendario_Inicio_Fiscal", 1),

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
            Funcao =
                (Data as date, Cultura as text) =>
                    not List.Contains({0, 6}, Date.DayOfWeek(Data))
                    and
                    not fxFeriado(Data)[Existe]
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
            Funcao = (Data as date, Cultura as text) =>
                if Date.Month(Data) >= InicioFiscal
                then Date.Year(Data)
                else Date.Year(Data) - 1
        ],

        TRIMESTRE_FISCAL = [
            Tipo = Int64.Type,
            Categoria = "Fiscal",
            Funcao = (Data as date, Cultura as text) =>
                let MF = Number.Mod(Date.Month(Data) - InicioFiscal + 12, 12) + 1
                in Number.RoundUp(MF / 3)
        ],

        MES_FISCAL = [
            Tipo = Int64.Type,
            Categoria = "Fiscal",
            Funcao = (Data as date, Cultura as text) =>
                Number.Mod(Date.Month(Data) - InicioFiscal + 12, 12) + 1
        ],

        SEMANA_FISCAL = [
            Tipo = Int64.Type,
            Categoria = "Fiscal",
            Funcao = (Data as date, Cultura as text) =>
                let
                    AF = if Date.Month(Data) >= InicioFiscal then Date.Year(Data) else Date.Year(Data) - 1,
                    InicioAF = #date(AF, InicioFiscal, 1)
                in
                    Number.IntegerDivide(Duration.Days(Data - InicioAF), 7) + 1
        ],

        FERIADO = [
            Tipo = type logical,
            Categoria = "Feriados",
            Funcao =
                (Data as date, Cultura as text) =>
                    fxFeriado(Data)[Existe]
        ],

        NOME_FERIADO = [
            Tipo = type text,
            Categoria = "Feriados",
            Funcao =
                (Data as date, Cultura as text) =>
                    fxFeriado(Data)[Nome]
        ],

        CODIGO_FERIADO = [
            Tipo = type text,
            Categoria = "Feriados",
            Funcao =
                (Data as date, Cultura as text) =>
                    fxFeriado(Data)[Código]
        ],

        TIPO_FERIADO = [
            Tipo = type text,
            Categoria = "Feriados",
            Funcao =
                (Data as date, Cultura as text) =>
                    fxFeriado(Data)[Tipo]
        ],

        EXPEDIENTE = [
            Tipo = type text,
            Categoria = "Feriados",
            Funcao =
                (Data as date, Cultura as text) =>
                    fxFeriado(Data)[Expediente]
        ],

        TIMESTAMP = [
            Tipo = type datetime,
            Categoria = "Base",
            Funcao = (Data as date, Cultura as text) => DateTime.From(Data)
        ]
    ],

    Tabela =
        Table.AddColumn(
            Table.Buffer(stgParametrosCalendario),
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
                    Nome = [Coluna],
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

shared fxFiltrarArquivos = (
    Tabela as table,
    FormatoArquivo as text,
    optional NomeArquivo as nullable text
)
as table =>

let
    Formato =
        Text.Upper(
            FormatoArquivo
        ),

    Extensoes =
        try
            Record.Field(
                cfgFormatosArquivos,
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

    ResultadoFormato =
        Table.SelectRows(
            Tabela,
            each
                List.Contains(
                    Extensoes,
                    Text.Lower([Extension])
                )
        ),

    Resultado =
        if NomeArquivo = null then
            ResultadoFormato
        else
            Table.SelectRows(
                ResultadoFormato,
                each [Name] = NomeArquivo
            )

in
    Resultado;

shared srcREST = let
    FonteDados =
        Text.Upper(
            fxParametro("Dados_Fonte")
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
            fxParametro("Dados_Fonte")
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

shared srcDados = let
    Resultado =
        fxDados()

in
    Resultado;
shared parTabelaParametrosFormatosArquivos = "tbParametrosFormatosArquivos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];
shared parTabelaParametrosTratamentos = "tbParametrosTratamentos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];
shared parTabelaParametrosValidacoes = "tbParametrosValidacoes" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];
shared parTabelaParametrosSeveridades = "tbParametrosSeveridades" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaParametrosCalendario = "tbParametrosCalendario" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaParametrosFeriados = "tbParametrosFeriados" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaCategoriasConsultasPQ = "tbSobreCategoriasConsultasPQ" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaParametrosBooleanos = "tbParametrosBooleanos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaParametrosTipos = "tbParametrosTipos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];
shared parTabelaSchema = "tbSchema" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];
shared parTabelaClientes = "tbClientes" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];
shared parTabelaProdutos = "tbProdutos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];
shared parTabelaVendas = "tbVendas" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared nrmProdutos = let
    Fonte = qaProdutos,
    Validada = fxQaFiltrarPorStatus(Fonte, "OK"),
    Normalizada = fxNrmAplicar(Validada, parTabelaProdutos)
in
    Normalizada;

shared nrmVendas = let
    Fonte = qaVendas,
    Validada = fxQaFiltrarPorStatus(Fonte, "OK"),
    Normalizada = fxNrmAplicar(Validada, parTabelaVendas),
    ValorTotal = Table.AddColumn(Normalizada, "ValorTotal", each [Quantidade] * [ValorUnitário], type number)
in
    ValorTotal;

shared nrmDados = let
    Fonte = qaDados,
    Validada = fxQaFiltrarPorStatus(Fonte, "OK"),
    Normalizada = fxNrmAplicar(Validada, "nrmDados")
in
    Normalizada;

shared dimProdutos = let
    Fonte =
        nrmProdutos,

    Chaves =
        Table.AddIndexColumn(
            Fonte,
            "IDProduto",
            1,
            1,
            Int64.Type
        ),

    Reordenada =
        Table.ReorderColumns(
            Chaves,
            {"IDProduto", "Código", "Descrição", "Categoria", "PreçoLista"}
        ),

    Buffer =
        Table.Buffer(
            Reordenada
        )
in
    Buffer;

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
        Cultura ?? "pt-BR",

    Atributos =
        List.Buffer(
            Table.ToRecords(
                Record.ToTable(cfgCalendarioAtributos)
            )
        ),

    Nomes =
        List.Transform(
            Atributos,
            each [Value][Nome]
        ),

    ComAtributos =
        Table.AddColumn(
            Calendario,
            "_Attrs",
            each
                let
                    D = [Data]
                in
                    Record.FromList(
                        List.Transform(
                            Atributos,
                            (a) => a[Value][Funcao](D, Idioma)
                        ),
                        Nomes
                    ),
            type record
        ),

    Expandido =
        Table.ExpandRecordColumn(
            ComAtributos,
            "_Attrs",
            Nomes
        ),

    Tipos =
        List.Transform(
            Atributos,
            each {[Value][Nome], [Value][Tipo]}
        ),

    Resultado =
        Table.TransformColumnTypes(
            Expandido,
            Tipos
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

shared dimCalendario = // Dimensão calendário gerada com todos os atributos ativos configurados em tbParametrosCalendario.
let
    CalendarioBase =
        fxCalendarioBase(
            cfgCalendario[DataInicial],
            cfgCalendario[DataFinal]
        ),

    Calendario =
        fxCalendario(
            CalendarioBase
        ),

    ComChave =
        Table.AddIndexColumn(Calendario, "IDData", 1, 1, Int64.Type),

    Reordenada =
        Table.ReorderColumns(
            ComChave,
            List.Combine({{"IDData", "Data"}, List.RemoveItems(Table.ColumnNames(ComChave), {"IDData", "Data"})})
        ),
    
    Buffer = Table.Buffer(Reordenada)

in
    Buffer;
shared fxTratamentoTrim = (valor as any, optional parametros as nullable any) as any =>

let
    // Funções de tratamento de texto, número e normalização — chamadas via srcOperadores pelo pipeline.
    Resultado =
        if valor = null then null else Text.Trim(Text.From(valor))
in
    Resultado;
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
    if valor = null then null else Text.Select(Text.From(valor), {"0".."9"});
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

// Constrói record de ocorrência de erro/aviso para retorno das funções de validação QA.
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

shared fxValidacaoList = (
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
                    "LIST",
                    contexto,
                    valor,
                    valor,
                    "Valor não pertence à lista permitida.",
                    [Permitidos = parametros]
                )
            }

in
    [Valor = valor, Ocorrencias = Ocorrencias]
;
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
                Texto = Text.Trim(Text.From(valor)),
                Partes = Text.Split(Texto, "@"),
                Local   = if List.Count(Partes) = 2 then Partes{0} else "",
                Dominio = if List.Count(Partes) = 2 then Partes{1} else "",
                UltimoPonto = Text.PositionOf(Dominio, ".", Occurrence.Last),
                TLD = if UltimoPonto >= 0 then Text.End(Dominio, Text.Length(Dominio) - UltimoPonto - 1) else ""
            in
                List.Count(Partes) = 2
                and Text.Length(Local) > 0
                and Text.Length(Dominio) > 0
                and not Text.Contains(Texto, " ")
                and UltimoPonto > 0
                and Text.Length(TLD) >= 2,

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
                            "CEPVAL",
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
                    "CPFVAL",
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
                    "CNPJVAL",
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
    tabela as nullable table,
    coluna as text
) as nullable record =>

let
    AnoAtual =
        Date.Year(Date.From(DateTime.LocalNow())),

    Resultado =
        if tabela = null then
            [
                DataInicial = #date(AnoAtual, 1, 1),
                DataFinal = #date(AnoAtual, 12, 31)
            ]
        else
            let
                SchemaTipo =
                    try Type.TableColumn(
                        Value.Type(tabela),
                        coluna
                    )
                    otherwise type any,

                ValoresOrigem =
                    Table.Column(
                        tabela,
                        coluna
                    ),

                Datas =
                    if
                        SchemaTipo = type date
                        or SchemaTipo = type datetime
                        or SchemaTipo = type datetimezone
                    then
                        List.Buffer(
                            List.RemoveNulls(
                                List.Transform(
                                    ValoresOrigem,
                                    each
                                        try Date.From(_)
                                        otherwise null
                                )
                            )
                        )
                    else
                        let
                            NaoNulos =
                                List.Select(
                                    ValoresOrigem,
                                    each _ <> null and _ <> ""
                                ),

                            Convertidos =
                                List.Transform(
                                    NaoNulos,
                                    (v) =>
                                        try Date.From(v)
                                        otherwise null
                                )
                        in
                            List.Buffer(
                                List.RemoveNulls(Convertidos)
                            )
            in
                if List.IsEmpty(Datas) then
                    null
                else
                    [
                        DataInicial = List.Min(Datas),
                        DataFinal = List.Max(Datas)
                    ]
in
    Resultado;
shared fxValidacaoRequired = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let
    Vazio =
        if valor = null then true
        else if valor is text then Text.Trim(valor) = ""
        else if valor is list then List.IsEmpty(valor)
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

shared fxNrmAplicar = 
//==============================================================================
// NORMALIZAÇÃO - NORMALIZE (ETL Layer 4)
// Responsabilidade: Estrutura de dados, deduplicação, enriquecimento
//==============================================================================
(
    tabela as table,
    optional Schema as nullable text
) as table =>
    let
        // Deduplica por ChavesNegocio do schema (Chave=VERDADEIRO em tbSchema); sem chaves definidas, retorna inalterada.
        Chaves = fxPipeline(Schema)[ChavesNegocio]
    in
        if List.IsEmpty(Chaves) then tabela
        else Table.Distinct(tabela, Chaves);

shared parObjetosPowerQueryIgnorados = // Queries do próprio catálogo — excluídas de stgTabelasPowerQuery para evitar auto-referência.
let
    Nomes =
        stgObjetosPowerQuery[Nome]
in
    List.Buffer(
        List.Select(
            Nomes,
            (Nome as text) =>
                Text.Contains(Nome, "Sections")
                or Text.Contains(Nome, "Objetos")
                or Text.Contains(Nome, "Tabelas")
                or Text.Contains(Nome, "Workbook")
                or Text.Contains(Nome, "diag")
                or Text.Contains(Nome, "tst")
        )
    );

shared fxStgPreparar = 
//==============================================================================
// STAGING - STAGE (ETL Layer 1)
// Responsabilidade: Apenas estruturação e normalização básica
//==============================================================================
(
    tabela as table,
    optional ignorarColunas as nullable list
)
as table =>

let
    // Remove linhas vazias, normaliza nomes de colunas e valida estrutura (lança erro em nomes vazios ou duplicados).
    // Configuração
    ColunasIgnoradas = ignorarColunas ?? {},
    ColunasTabela = Table.ColumnNames(tabela),
    ColunasEfetivas = List.Difference(ColunasTabela, ColunasIgnoradas),

    // 1. Validar estrutura — referenciado em cada passo subsequente para forçar avaliação lazy
    TabelaValidada =
        if List.IsEmpty(ColunasTabela) then
            error Error.Record(
                "Tabela inválida",
                "A tabela não possui colunas.",
                [Tabela = tabela]
            )
        else
            tabela,

    // 2. Remover linhas completamente vazias
    // Otimização: List.MatchesAny faz short-circuit na primeira célula preenchida,
    // eliminando alocações de sub-record e lista temporária para cada linha.
    TabelaSemLinhasVazias =
        if List.IsEmpty(ColunasEfetivas) then
            TabelaValidada
        else
            Table.SelectRows(
                TabelaValidada,
                each
                    let Linha = _ in
                    List.MatchesAny(
                        ColunasEfetivas,
                        (Col) =>
                            let v = Record.FieldOrDefault(Linha, Col, null) in
                            v <> null and v <> ""
                    )
            ),

    // 3. Normalizar nomes de colunas (trim)
    ColunasNormalizadas = 
        List.ReplaceMatchingItems(
            ColunasTabela,
            List.Transform(
                ColunasEfetivas,
                each {_, Text.Trim(_)}
            )
        ),

    // 4. Validar nomes normalizados — erros lançados aqui são referenciados em Resultado
    ColunasValidadas =
        if List.AnyTrue(List.Transform(ColunasNormalizadas, each _ = "")) then
            error Error.Record(
                "Tabela inválida",
                "A tabela possui colunas com nome vazio.",
                [Colunas = ColunasNormalizadas]
            )
        else if List.Count(ColunasNormalizadas) <> List.Count(List.Distinct(ColunasNormalizadas)) then
            error Error.Record(
                "Tabela inválida",
                "A tabela possui nomes de colunas duplicados.",
                [Colunas = ColunasNormalizadas]
            )
        else
            ColunasNormalizadas,

    // 5. Aplicar renomeação
    Resultado = 
        Table.RenameColumns(
            TabelaSemLinhasVazias,
            List.Zip({ColunasTabela, ColunasValidadas}),
            MissingField.Ignore
        )

in
    Resultado;
shared fxStgAplicar = (
    Tabela as table,
    optional Schema as nullable text,
    optional ignorarColunas as nullable list
) as table =>

let
    // ========================================================================
    // FASE 1: PRÉ-COMPILAÇÃO E CACHE
    // ========================================================================
    // Evita recomputação de metadados em cada linha
    
    Preparada = fxStgPreparar(Tabela, ignorarColunas),
    
    Pipeline = fxPipeline(Schema),
    TiposPorColuna = Pipeline[TiposPorColuna],
    Ordem = Pipeline[Ordem],
    
    // Cache de nomes de colunas com tipos
    ColunasTipos = List.Buffer(Record.FieldNames(TiposPorColuna)),
    
    // ========================================================================
    // FASE 2: SEPARAÇÃO DE TIPOS POR ESTRATÉGIA DE CONVERSÃO
    // ========================================================================
    // Agrupa colunas por tipo de tratamento necessário
    
    // Tipos nativos que o Power Query converte diretamente
    ColunasNativas = List.Buffer(
        List.Select(
            ColunasTipos,
            (NomeColuna) =>
                let
                    TipoDestino = Record.Field(TiposPorColuna, NomeColuna)
                in
                    TipoDestino <> type any
                    and TipoDestino <> type list
                    and TipoDestino <> type logical
        )
    ),
    
    // Tipos especiais que precisam de conversão customizada
    ColunasEspeciais = List.Buffer(
        List.Select(
            ColunasTipos,
            (NomeColuna) =>
                let
                    TipoDestino = Record.Field(TiposPorColuna, NomeColuna)
                in
                    TipoDestino = type list or TipoDestino = type logical
        )
    ),
    
    // ========================================================================
    // FASE 3: PRÉ-COMPILAÇÃO DE DEFINIÇÕES
    // ========================================================================
    // Cria estruturas de dados otimizadas para lookup rápido
    
    // Mapa de tipo textual para conversão
    MapaTipos = Record.FromList(
        List.Transform(
            ColunasTipos,
            (Col) => [
                Coluna = Col,
                Tipo = Record.Field(TiposPorColuna, Col),
                TipoTexto = fxTipoParaTexto(Record.Field(TiposPorColuna, Col))
            ]
        ),
        ColunasTipos
    ),
    
    // Definições nativas pré-compiladas com contextos de erro
    DefinicoesNativasCompiladas = List.Buffer(
        List.Transform(
            ColunasNativas,
            (Col) =>
                let
                    Info = Record.Field(MapaTipos, Col)
                in
                    [
                        Coluna = Col,
                        Tipo = Info[Tipo],
                        Contexto = [
                            Coluna = Col,
                            Tipo = Info[Tipo],
                            Operador = [Severidade = "ERRO"]
                        ],
                        Mensagem = "Valor inválido removido. Esperado tipo: " & Info[TipoTexto] & ".",
                        Detalhes = [
                            TipoEsperado = Info[TipoTexto],
                            ValorRemovido = true
                        ]
                    ]
        )
    ),
    
    // ========================================================================
    // FASE 4: APLICAÇÃO DE TIPOS NATIVOS
    // ========================================================================
    // Usa TransformColumnTypes nativo do Power Query (otimizado)
    
    TiposNativos = List.Buffer(
        List.Transform(
            ColunasNativas,
            (Col) => {Col, Record.Field(MapaTipos, Col)[Tipo]}
        )
    ),
    
    TiposAplicados = 
        if List.IsEmpty(TiposNativos) then Preparada
        else Table.TransformColumnTypes(
            Preparada,
            TiposNativos,
            parCulturaBootstrap
        ),
    
    // ========================================================================
    // FASE 5: CONVERSÃO DE TIPOS ESPECIAIS (OTIMIZADA)
    // ========================================================================
    // Usa conversores customizados apenas onde necessário
    
    TransformacoesEspeciais = List.Transform(
        ColunasEspeciais,
        (Col) =>
            let
                Info = Record.Field(MapaTipos, Col),
                TipoDestino = Info[Tipo],
                TipoTexto = Info[TipoTexto],
                
                // Cache do conversor para evitar lookup repetido
                Conversor = (v) => fxConversor(v, TipoDestino, parCulturaBootstrap)
            in
                {
                    Col,
                    (ValorCelula) =>
                        if ValorCelula = null then null
                        else
                            let
                                Tentativa = try Conversor(ValorCelula)
                            in
                                if Tentativa[HasError] then
                                    [
                                        __STG_ERRO = true,
                                        Ocorrencias = fxSchemaOcorrencia(
                                            "TIPO_INVALIDO",
                                            [
                                                Coluna = Col,
                                                Tipo = TipoDestino,
                                                Operador = [Severidade = "ERRO"]
                                            ],
                                            null,
                                            null,
                                            "Valor inválido removido. Esperado tipo: " & TipoTexto & ".",
                                            [
                                                TipoEsperado = TipoTexto,
                                                ValorRemovido = true
                                            ]
                                        )
                                    ]
                                else
                                    Tentativa[Value],
                    type any
                }
    ),
    
    EspeciaisAplicados =
        if List.IsEmpty(TransformacoesEspeciais) then TiposAplicados
        else Table.TransformColumns(
            TiposAplicados,
            TransformacoesEspeciais,
            null,
            MissingField.Ignore
        ),
    
    // ========================================================================
    // FASE 6: DETECÇÃO DE ERROS OTIMIZADA (SINGLE-PASS)
    // ========================================================================
    // Uma única passagem pela tabela para detectar todos os erros
    
    // Pré-calcula lista de substituição de erros nativos
    SubstituicoesErroNativos = List.Buffer(
        List.Transform(ColunasNativas, (Col) => {Col, null})
    ),
    
    // Adiciona coluna de ocorrências usando List.MatchesAny para short-circuit
    ComOcorrencias = Table.AddColumn(
        EspeciaisAplicados,
        "_STG_Ocorrencias",
        (LinhaAtual) =>
            let
                // OTIMIZAÇÃO 1: Usa List.MatchesAny com early exit
                TemErroNativo = List.MatchesAny(
                    DefinicoesNativasCompiladas,
                    (Def) =>
                        let
                            Tentativa = try Record.Field(LinhaAtual, Def[Coluna])
                        in
                            Tentativa[HasError]
                ),
                
                // OTIMIZAÇÃO 2: Só processa erros nativos se existirem
                OcorrenciasNativas =
                    if not TemErroNativo then {}
                    else List.RemoveNulls(
                        List.Transform(
                            DefinicoesNativasCompiladas,
                            (Def) =>
                                let
                                    Tentativa = try Record.Field(LinhaAtual, Def[Coluna])
                                in
                                    if Tentativa[HasError] then
                                        fxSchemaOcorrencia(
                                            "TIPO_INVALIDO",
                                            Def[Contexto],
                                            null,
                                            null,
                                            Def[Mensagem],
                                            Def[Detalhes]
                                        )
                                    else null
                        )
                    ),
                
                // OTIMIZAÇÃO 3: Usa List.MatchesAny para detectar erros especiais
                TemErroEspecial = List.MatchesAny(
                    ColunasEspeciais,
                    (Col) =>
                        let
                            v = Record.FieldOrDefault(LinhaAtual, Col, null)
                        in
                            Value.Is(v, type record)
                            and Record.FieldOrDefault(v, "__STG_ERRO", false) = true
                ),
                
                // OTIMIZAÇÃO 4: Só processa erros especiais se existirem
                OcorrenciasEspeciais =
                    if not TemErroEspecial then {}
                    else List.RemoveNulls(
                        List.Transform(
                            ColunasEspeciais,
                            (Col) =>
                                let
                                    v = Record.FieldOrDefault(LinhaAtual, Col, null)
                                in
                                    if Value.Is(v, type record)
                                        and Record.FieldOrDefault(v, "__STG_ERRO", false) = true
                                    then Record.FieldOrDefault(v, "Ocorrencias", null)
                                    else null
                        )
                    ),
                
                // OTIMIZAÇÃO 5: Só combina listas se houver ocorrências
                TodasOcorrencias =
                    if not TemErroNativo and not TemErroEspecial then null
                    else
                        let
                            Combinadas = List.Combine({OcorrenciasNativas, OcorrenciasEspeciais})
                        in
                            if List.IsEmpty(Combinadas) then null else Combinadas
            in
                TodasOcorrencias,
        type any
    ),
    
    // ========================================================================
    // FASE 7: SUBSTITUIÇÃO DE ERROS NATIVOS (BATCH)
    // ========================================================================
    // Usa ReplaceErrorValues uma única vez para todos os erros
    
    NativosSemErros =
        if List.IsEmpty(ColunasNativas) then ComOcorrencias
        else Table.ReplaceErrorValues(ComOcorrencias, SubstituicoesErroNativos),
    
    // ========================================================================
    // FASE 8: LIMPEZA DE TIPOS ESPECIAIS
    // ========================================================================
    // Remove marcadores de erro e aplica tipos finais
    
    TransformacoesEspeciaisFinais = List.Transform(
        ColunasEspeciais,
        (Col) =>
            let
                Info = Record.Field(MapaTipos, Col)
            in
                {
                    Col,
                    (ValorCelula) =>
                        if Value.Is(ValorCelula, type record)
                            and Record.FieldOrDefault(ValorCelula, "__STG_ERRO", false) = true
                        then null
                        else ValorCelula,
                    Info[Tipo]
                }
    ),
    
    ComTipos =
        if List.IsEmpty(TransformacoesEspeciaisFinais) then NativosSemErros
        else Table.TransformColumns(
            NativosSemErros,
            TransformacoesEspeciaisFinais,
            null,
            MissingField.Ignore
        ),
    
    // ========================================================================
    // FASE 9: REORDENAÇÃO FINAL
    // ========================================================================
    
    Resultado =
        if List.IsEmpty(Ordem) then ComTipos
        else Table.ReorderColumns(ComTipos, Ordem, MissingField.Ignore)

in
    Resultado;

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

shared fxObjetoIdentificarPeloNome = (
    objeto as text,
    source as text
)
as record =>

let
    // Infere Kind e categoria de uma query pelo prefixo do nome (ex: "stg" → Staging).


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
                Categoria = "Excel",
                Objetivo = null,
                Saída = "Tabela"
            ]
        else
            Record.FieldOrDefault(
                cfgCategoriasPowerQuery,
                Prefixo,
                [
                    Categoria = "Qualquer",
                    Objetivo = null,
                    Saída = "Qualquer"
                ]
            ),

//--------------------------------------------------------------------------
// Definição do tipo
//--------------------------------------------------------------------------

    DefinicaoTipo = Record.FieldOrDefault(cfgTiposObjetosPorNome, Categoria[Saída], null),

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
                        Categoria[Categoria]
                    }

                )

            )

        else

            [
                Prefix = Prefixo,
                Kind = DefinicaoTipo[Kind],
                Tipo = DefinicaoTipo[Nome],
                Type = DefinicaoTipo[Type],
                Categoria = Categoria[Categoria]
            ]

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

shared cfgOperadores = 
// Record {Código → operador compilado} para lookup O(1) durante a compilação do pipeline.
let

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

shared cfgOperadoresTratamentoPadrao = 
// Tratamentos aplicados automaticamente a cada coluna de acordo com o tipo de dado.
let

    Fonte =
        Table.SelectRows(
            stgOperadores,
            each
                [Ativo] and
                [Padrão] and
                [Categoria] = "Tratamento"
        ),

    OperadoresAny =
        List.Buffer(
            Table.SelectRows(
                Fonte,
                each Value.Equals([TipoEntrada], type any)
            )[Código]
        ),

    TiposEspecificos =
        Table.SelectRows(
            Fonte,
            each not Value.Equals([TipoEntrada], type any)
        ),

    Agrupar =
        Table.Group(
            TiposEspecificos,
            {"TipoEntrada"},
            {
                {
                    "Operadores",
                    each
                        List.Buffer(
                            List.Combine(
                                {
                                    OperadoresAny,
                                    [Código]
                                }
                            )
                        ),
                    type list
                }
            }
        ),

    Resultado =
        List.Buffer(
            Table.ToRecords(
                Table.RenameColumns(
                    Agrupar,
                    {
                        {"TipoEntrada", "Tipo"}
                    }
                )
            )
        )
in
    Resultado;

shared cfgOperadoresTratamentoPadraoMap = Record.FromList(
    List.Transform(cfgOperadoresTratamentoPadrao, each [Operadores]),
    List.Transform(cfgOperadoresTratamentoPadrao, each fxTipoParaTexto([Tipo]))
);

shared cfgOperadoresValidacaoPadrao = let

    Fonte =
        Table.SelectRows(
            stgOperadores,
            each
                [Ativo] and
                [Padrão] and
                [Categoria] = "Validação" and
                [Código] <> "REQUIRED"
        ),

    OperadoresAny =
        List.Buffer(
            Table.SelectRows(
                Fonte,
                each Value.Equals([TipoEntrada], type any)
            )[Código]
        ),

    TiposEspecificos =
        Table.SelectRows(
            Fonte,
            each not Value.Equals([TipoEntrada], type any)
        ),

    Agrupar =
        Table.Group(
            TiposEspecificos,
            {"TipoEntrada"},
            {
                {
                    "Operadores",
                    each
                        List.Buffer(
                            List.Combine(
                                {
                                    OperadoresAny,
                                    [Código]
                                }
                            )
                        ),
                    type list
                }
            }
        ),

    Resultado =
        List.Buffer(
            Table.ToRecords(
                Table.RenameColumns(
                    Agrupar,
                    {
                        {"TipoEntrada", "Tipo"}
                    }
                )
            )
        )
in
    Resultado;

shared cfgOperadoresValidacaoPadraoMap = Record.FromList(
    List.Transform(cfgOperadoresValidacaoPadrao, each [Operadores]),
    List.Transform(cfgOperadoresValidacaoPadrao, each fxTipoParaTexto([Tipo]))
);

shared cfgSchema = 
// Record {Tabela → {Coluna → definição}} derivado de stgSchema; base de cfgPipeline.
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
    Resultado;

shared cfgPipeline = // Record lazy {Tabela → pipeline compilado}; campos: Ordem, TiposPorColuna, TratamentosPorColuna, ValidaçõesPorColuna, ChavesNegocio.
let
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

// Extrai {Código, Parâmetros} de um operador no formato "CODIGO" ou "CODIGO(p1,p2)".
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
                                cfgCultura[SeparadorParametros]
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

shared fxPipelineCompilarOperadores = // Resolve lista de strings de operadores (ex: ["TRIM","UPPER"]) em lista de operadores compilados prontos para execução.
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

shared fxPipelineCompilar = // Compila o schema completo de uma tabela em pipeline executável.
(
    Schema as record
)
as record =>

let

    //--------------------------------------------------------------------------
    // Colunas e ordem (bufferizadas — usadas em várias etapas)
    //--------------------------------------------------------------------------

    Colunas =
        List.Buffer(
            Record.FieldNames(Schema)
        ),

    Ordem = Colunas,

    //--------------------------------------------------------------------------
    // REQUIRED bufferizado — evita reconstrução do record a cada coluna obrigatória
    //--------------------------------------------------------------------------

    OperadorRequired =
        Record.Combine(
            {
                cfgOperadores[REQUIRED],
                [Parâmetros = null]
            }
        ),

    //--------------------------------------------------------------------------
    // Cache de operadores padrão POR TIPO
    // ------------------------------------------------------------------------
    // fxOperadoresPadrao faz List.Select sobre as tabelas de configuração.
    // Como muitas colunas compartilham o mesmo Tipo (ex.: vários "type text"),
    // memorizamos o resultado por chave textual do tipo. Assim, dentro de uma
    // única compilação de tabela, cada tipo distinto resolve os padrões só 1x.
    //--------------------------------------------------------------------------

    TiposDistintos =
        List.Buffer(
            List.Distinct(
                List.Transform(
                    Colunas,
                    each Record.Field(Record.Field(Schema, _), "Tipo")
                ),
                // comparador por representação textual do tipo
                (t) => fxTipoParaTexto(t)
            )
        ),

    PadroesPorChaveTipo =
        Record.FromList(
            List.Transform(
                TiposDistintos,
                each fxOperadoresPadrao(_)
            ),
            List.Transform(
                TiposDistintos,
                each fxTipoParaTexto(_)
            )
        ),

    ObterPadroes =
        (Tipo as type) as record =>
            Record.Field(
                PadroesPorChaveTipo,
                fxTipoParaTexto(Tipo)
            ),

    //--------------------------------------------------------------------------
    // PASSAGEM ÚNICA
    // ------------------------------------------------------------------------
    // Para cada coluna extraímos a Definicao UMA vez e computamos, de uma só
    // varredura, todos os artefatos necessários. O resultado é uma lista de
    // records intermediários que depois é "fatiada" nos records de saída.
    //--------------------------------------------------------------------------

    Compilado =
        List.Transform(
            Colunas,
            (Coluna) =>
                let
                    Definicao = Record.Field(Schema, Coluna),
                    Tipo      = Definicao[Tipo],
                    Padroes   = ObterPadroes(Tipo),

                    Tratamentos =
                        fxPipelineCompilarOperadores(
                            List.Combine(
                                {
                                    Padroes[Tratamentos],
                                    Definicao[Tratamentos] ?? {}
                                }
                            )
                        ),

                    ValidacoesBase =
                        fxPipelineCompilarOperadores(
                            List.Combine(
                                {
                                    Padroes[Validações],
                                    Definicao[Validações] ?? {}
                                }
                            )
                        ) ?? {},

                    Required =
                        if Definicao[Obrigatório] and OperadorRequired[Padrão] then
                            {OperadorRequired}
                        else
                            {},

                    Validacoes =
                        List.Combine({Required, ValidacoesBase}),

                    EhChave =
                        Record.FieldOrDefault(Definicao, "Chave", null) = true
                in
                    [
                        Coluna       = Coluna,
                        Tipo         = Tipo,
                        Tratamentos  = if List.IsEmpty(Tratamentos) then null else Tratamentos,
                        Validacoes   = if List.IsEmpty(Validacoes)  then null else Validacoes,
                        Chave        = EhChave
                    ]
        ),

    CompiladoBuffer = List.Buffer(Compilado),

    //--------------------------------------------------------------------------
    // Distribuição dos resultados nos records de saída
    //--------------------------------------------------------------------------

    TiposPorColuna =
        Record.FromList(
            List.Transform(CompiladoBuffer, each [Tipo]),
            List.Transform(CompiladoBuffer, each [Coluna])
        ),

    ComTratamentos =
        List.Select(CompiladoBuffer, each [Tratamentos] <> null),

    TratamentosPorColuna =
        Record.FromList(
            List.Transform(ComTratamentos, each [Tratamentos]),
            List.Transform(ComTratamentos, each [Coluna])
        ),

    ComValidacoes =
        List.Select(CompiladoBuffer, each [Validacoes] <> null),

    ValidacoesPorColuna =
        Record.FromList(
            List.Transform(ComValidacoes, each [Validacoes]),
            List.Transform(ComValidacoes, each [Coluna])
        ),

    ChavesNegocio =
        List.Buffer(
            List.Transform(
                List.Select(CompiladoBuffer, each [Chave]),
                each [Coluna]
            )
        ),

    //--------------------------------------------------------------------------
    // Resultado (contrato idêntico ao original)
    //--------------------------------------------------------------------------

    Resultado =
        [
            Ordem                = Ordem,
            TiposPorColuna       = TiposPorColuna,
            TratamentosPorColuna = TratamentosPorColuna,
            ValidaçõesPorColuna  = ValidacoesPorColuna,
            ChavesNegocio        = ChavesNegocio
        ]

in
    Resultado;

shared fxPipeline = 
// Retorna pipeline compilado para o schema informado; retorna pipeline vazio se Schema não existir.
(
    optional Schema as nullable text
)
as record =>

let
    PipelineVazio =
        [
            Ordem = {},
            TiposPorColuna = [],
            TratamentosPorColuna = [],
            ValidaçõesPorColuna = [],
            ChavesNegocio = {}
        ],

    Resultado =
        if
            Schema = null
            or Text.Trim(Schema) = ""
        then
            PipelineVazio
        else
            Record.FieldOrDefault(
                cfgPipeline,
                Schema,
                PipelineVazio
            )
in
    Resultado;
shared srcRESTHeaders =
// Cabeçalhos HTTP para requisições REST. Credenciais via parâmetros REST_Token e REST_ApiKey.
// Credenciais lidas de parâmetros — nunca hardcode tokens aqui.
[
        Accept = "application/json",
        Authorization = "Bearer " & fxParametro("REST_Token", ""),
        #"x-api-key" = fxParametro("REST_ApiKey", ""),
        #"User-Agent" = "Power Query",
        #"Content-Type" = "application/json"
];

shared fxTrnAplicar = 
//==============================================================================
// TRATAMENTOS - TRANSFORM (ETL Layer 2)
// Responsabilidade: Tratamentos e limpeza de dados
//==============================================================================
(
    Tabela as table,
    optional Schema as nullable text
)
as table =>

let
    // Aplica os tratamentos do schema à tabela em passagem única por coluna.
    // Obter pipeline compilado
    Pipeline =
        fxPipeline(Schema),

    TratamentosPorColuna =
        Pipeline[TratamentosPorColuna],

    TiposPorColuna =
        Pipeline[TiposPorColuna],

    // Compilar transformações apenas para as colunas com tratamentos configurados
    ColunasTratamento =
        List.Buffer(Record.FieldNames(TratamentosPorColuna)),

    Transformacoes =
        List.RemoveNulls(
            List.Transform(
                ColunasTratamento,
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

shared fxTrnCompilarTratamentosPorColuna = 
// Cria função que aplica todos os tratamentos em cadeia a um valor; resultado é compilado uma vez por coluna.
(
    Operadores as list
)
as function =>

let
    // Pré-extrai a função e parâmetros dos tratamentos na compilação,
    // eliminando lookups dinâmicos de record por linha.
    TratamentosCompilados =
        List.Buffer(
            List.Transform(
                Operadores ?? {},
                (Operador) =>
                    [
                        Funcao = Operador[Função],
                        Parametros = Record.FieldOrDefault(Operador, "Parâmetros", null)
                    ]
            )
        )

in
    if List.IsEmpty(TratamentosCompilados) then
        (valor) => valor
    else
        (valor) =>
            List.Accumulate(
                TratamentosCompilados,
                valor,
                (Estado, Op) =>
                    Op[Funcao](
                        Estado,
                        Op[Parametros]
                    )
            );

shared fxQaValidar = 
//==============================================================================
// VALIDAÇÕES - QUALITY ASSURANCE (ETL Layer 3)
// Responsabilidade: Validações estruturais, semânticas e de negócio
//==============================================================================
(
    Tabela as table,
    optional Schema as nullable text
)
as table =>

let
    // Valida a tabela conforme o schema. Adiciona _QA_Status (OK/AVISO/ERRO) e _QA_Ocorrencias sem remover dados.
    // Obter pipeline
    Pipeline =
        fxPipeline(Schema),

    ValidacoesPorColuna =
        Pipeline[ValidaçõesPorColuna],

    TiposPorColuna =
        Pipeline[TiposPorColuna],

    ColunasValidacao =
        Record.FieldNames(
            ValidacoesPorColuna
        ),

    Resultado =
        let
            // Pré-compilação otimizada dos validadores
            ValidadoresPorColuna =
                Record.FromList(
                    List.Transform(
                        ColunasValidacao,
                        (Coluna) =>
                            fxQaCompilarValidacoesPorColuna(
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

            // Pré-compilação dos validadores e nomes das colunas como uma lista de registros.
            // Isso evita a necessidade de buscar dinamicamente o validador por nome (Record.Field)
            // de coluna em cada linha processada na tabela.
            ValidadoresInfo =
                List.Buffer(
                    List.Transform(
                        ColunasValidacao,
                        (Coluna) =>
                            [
                                Coluna = Coluna,
                                Validador = Record.Field(ValidadoresPorColuna, Coluna)
                            ]
                    )
                ),

            // Cache de severidades
            CacheSeveridades =
                let
                    Nomes =
                        Record.FieldNames(
                            cfgSeveridades
                        )
                in
                    Record.FromList(
                        List.Transform(
                            Nomes,
                            each
                                Record.FieldOrDefault(
                                    Record.Field(
                                        cfgSeveridades,
                                        _
                                    ),
                                    "Bloqueia",
                                    false
                                )
                        ),
                        Nomes
                    ),

            // Única passagem pela tabela
            TabelaComQA =
                Table.AddColumn(
                    Tabela,
                    "_QA",
                    each
                        let
                            Linha =
                                _,

                            OcorrenciasIniciais =
                                let
                                    x =
                                        Record.FieldOrDefault(
                                            Linha,
                                            "_STG_Ocorrencias",
                                            null
                                        )
                                in
                                    if x = null then
                                        {}
                                    else
                                        x,

                            Ocorrencias =
                                List.Accumulate(
                                    ValidadoresInfo,
                                    OcorrenciasIniciais,
                                    (Estado, Info) =>
                                        let
                                            Valor =
                                                Record.FieldOrDefault(
                                                    Linha,
                                                    Info[Coluna],
                                                    null
                                                ),

                                            ResultadoVal =
                                                Info[Validador](
                                                    Valor
                                                )
                                        in
                                            if
                                                ResultadoVal = null
                                                or List.IsEmpty(ResultadoVal)
                                            then
                                                Estado
                                            else
                                                Estado & ResultadoVal
                                ),

                            Status =
                                if List.IsEmpty(Ocorrencias) then

                                    "OK"

                                else

                                    let
                                        TemErro =
                                            List.MatchesAny(
                                                Ocorrencias,
                                                (o) =>
                                                    Record.FieldOrDefault(
                                                        CacheSeveridades,
                                                        o[Severidade],
                                                        false
                                                    )
                                            )
                                    in
                                        if TemErro then
                                            "ERRO"
                                        else
                                            "AVISO"
                        in
                            [
                                Status = Status,
                                Ocorrencias =
                                    if List.IsEmpty(Ocorrencias) then
                                        null
                                    else
                                        Ocorrencias
                            ],
                    type record
                ),

            ResultadoFinal =
                Table.ExpandRecordColumn(
                    TabelaComQA,
                    "_QA",
                    {
                        "Status",
                        "Ocorrencias"
                    },
                    {
                        "_QA_Status",
                        "_QA_Ocorrencias"
                    }
                ),

            ResultadoSemStgOcorrencias =
                Table.RemoveColumns(
                    ResultadoFinal,
                    {"_STG_Ocorrencias"},
                    MissingField.Ignore
                )
        in
            ResultadoSemStgOcorrencias

in
    Resultado;

shared fxQaFiltrarPorStatus = 
// Filtra por _QA_Status e remove colunas de controle QA (_QA_Status, _QA_Ocorrencias).
(
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

shared fxQaExtrairProblemas = 
// Retorna apenas os registros com _QA_Status diferente de OK (AVISO ou ERRO).
(
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

shared fxQaCompilarValidacoesPorColuna = 
// Cria função que acumula todas as ocorrências de validação de uma coluna; retorna null se sem erros.
(
    Operadores as list,
    Tipo as type,
    Coluna as text
)
as function =>

let
    // Pré-cria os contextos de validação e extrai os parâmetros na compilação,
    // evitando a recriação de records dinâmicos a cada linha da tabela.
    OperadoresCompilados =
        List.Buffer(
            List.Transform(
                Operadores ?? {},
                (Operador) =>
                    [
                        Funcao = Operador[Função],
                        Parametros = Record.FieldOrDefault(Operador, "Parâmetros", null),
                        Contexto = [
                            Coluna = Coluna,
                            Tipo = Tipo,
                            Operador = Operador
                        ]
                    ]
            )
        )

in
    if List.IsEmpty(OperadoresCompilados) then
        (valor) => null
    else
        (valor) =>
            let
                Ocorrencias =
                    List.Accumulate(
                        OperadoresCompilados,
                        {},
                        (Estado, Op) =>
                            let
                                Resultado =
                                    Op[Funcao](
                                        valor,
                                        Op[Parametros],
                                        Op[Contexto]
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

shared trnClientes = let
    Fonte = stgClientes,
    Transformada = fxTrnAplicar(Fonte, parTabelaClientes)
in
    Transformada;

shared qaClientes = let
    Fonte = trnClientes,
    qa = fxQaValidar(Fonte, parTabelaClientes)

    // Usar apenas dados válidos
    // Validos = fxQaFiltrarPorStatus(qa, "OK"),

    // Extrair problemas para auditoria
    // Problemas = fxQaExtrairProblemas(qa)

in
    qa;

shared trnProdutos = let
    Fonte = stgProdutos,
    Transformada = fxTrnAplicar(Fonte, parTabelaProdutos)
in
    Transformada;

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
    Transformada = fxTrnAplicar(Fonte, parTabelaVendas)
in
    Transformada;

shared trnDados = let
    Fonte = stgDados,
    Transformada = fxTrnAplicar(Fonte, "nrmDados")
in
    Transformada;

shared qaVendas = let
    Fonte = trnVendas,
    qa = fxQaValidar(Fonte, parTabelaVendas)

    // Usar apenas dados válidos
    // Validos = fxQaFiltrarPorStatus(qa, "OK"),

    // Extrair problemas para auditoria
    // Problemas = fxQaExtrairProblemas(qa)

in
    qa;

shared qaDados = let
    Fonte = trnDados,
    qa = fxQaValidar(Fonte, "nrmDados")
in
    qa;

shared fxTesteAsertar = 
// Executa uma assertiva com try; retorna {Teste, Status, Detalhes}.
(
    nome     as text,
    condicao as function,
    optional detalhes as nullable text
) as record =>

    let
        R = try condicao()
    in
        [
            Teste    = nome,
            Status   = if R[HasError] then "ERRO" else if R[Value] then "PASSOU" else "FALHOU",
            Detalhes = if R[HasError] then R[Error][Message] else if not R[Value] then (detalhes ?? "") else null
        ];

shared fxTesteExecutar = 
// Executa lista de assertivas e retorna tabela de resultados com linha de resumo.
(
    suite      as text,
    assertivas as list
) as table =>
    let
        Passaram  = List.Count(List.Select(assertivas, each [Status] = "PASSOU")),
        Total     = List.Count(assertivas),
        Resultado = Table.InsertRows(
            Table.FromRecords(
                assertivas,
                type table [Teste = text, Status = text, Detalhes = nullable text]
            ),
            0,
            {[
                Teste    = "── " & suite & " ──",
                Status   = Text.From(Passaram) & "/" & Text.From(Total) & " passaram",
                Detalhes = null
            ]}
        )
    in
        Resultado;

shared tstPipelineExecucao = let
    Fonte =
        srcClientes,

    STG =
        stgClientes,

    TRN =
        trnClientes,

    QA =
        qaClientes,

    NRM =
        nrmClientes,

    Resultado =
        fxPipelineDiagnostico(
            Fonte,
            STG,
            TRN,
            QA,
            NRM
        )
in
    Resultado;

shared tstTesteSTG = let
    Resultado = fxStgAplicar(srcClientes, "tbClientes")
in
    fxTesteExecutar("STG", {
        fxTesteAsertar("Retorna tabela",           () => Value.Is(Resultado, type table)),
        fxTesteAsertar("Tem ao menos 1 coluna",    () => Table.ColumnCount(Resultado) >= 1),
        fxTesteAsertar("Sem linhas vazias totais", () => not Table.IsEmpty(Resultado)),
        fxTesteAsertar("Sem coluna _QA_Status",    () => not Table.HasColumns(Resultado, "_QA_Status"))
    });

shared tstTesteTRN = let
    STG       = fxStgAplicar(srcClientes, "tbClientes"),
    Resultado = fxTrnAplicar(STG, "tbClientes")
in
    fxTesteExecutar("TRN", {
        fxTesteAsertar("Retorna tabela",           () => Value.Is(Resultado, type table)),
        fxTesteAsertar("Mesmas colunas que STG",   () => Table.ColumnCount(Resultado) = Table.ColumnCount(STG)),
        fxTesteAsertar("Mesma contagem de linhas", () => Table.RowCount(Resultado) = Table.RowCount(STG))
    });

shared tstTesteQA = let
    STG       = fxStgAplicar(srcClientes, "tbClientes"),
    TRN       = fxTrnAplicar(STG, "tbClientes"),
    Resultado = fxQaValidar(TRN, "tbClientes")
in
    fxTesteExecutar("QA", {
        fxTesteAsertar("Retorna tabela",                    () => Value.Is(Resultado, type table)),
        fxTesteAsertar("Tem coluna _QA_Status",             () => Table.HasColumns(Resultado, "_QA_Status")),
        fxTesteAsertar("Tem coluna _QA_Ocorrencias",        () => Table.HasColumns(Resultado, "_QA_Ocorrencias")),
        fxTesteAsertar("_QA_Status só tem valores válidos", () =>
            List.IsEmpty(
                List.Difference(
                    List.Distinct(List.RemoveNulls(Table.Column(Resultado, "_QA_Status"))),
                    {"OK", "AVISO", "ERRO"}
                )
            )
        )
    });

shared tstTesteNRM = let
    STG       = fxStgAplicar(srcClientes, "tbClientes"),
    TRN       = fxTrnAplicar(STG, "tbClientes"),
    QA        = fxQaValidar(TRN, "tbClientes"),
    Valido    = fxQaFiltrarPorStatus(QA, "OK"),
    Resultado = fxNrmAplicar(Valido, "tbClientes")
in
    fxTesteExecutar("NRM", {
        fxTesteAsertar("Retorna tabela",          () => Value.Is(Resultado, type table)),
        fxTesteAsertar("Sem coluna _QA_Status",   () => not Table.HasColumns(Resultado, "_QA_Status")),
        fxTesteAsertar("Linhas <= entrada QA OK", () => Table.RowCount(Resultado) <= Table.RowCount(Valido))
    });

shared tstTesteProblemasEncontrados = // Extrair problemas encontrados
let
    ProblemasDetectados = fxQaExtrairProblemas(qaClientes),
    _QA_OcorrenciasExpandido = Table.ExpandListColumn(ProblemasDetectados, "_QA_Ocorrencias"),
    _QA_OcorrenciasComMensagem = Table.ExpandRecordColumn(_QA_OcorrenciasExpandido, "_QA_Ocorrencias", {"Coluna", "Mensagem"}, {"Coluna", "Mensagem"})
in
    _QA_OcorrenciasComMensagem;

shared tstClientes_ComFramework = let
    Schema = "nrmDados",
    Fonte = srcDados,
    Preparacao = fxStgAplicar(Fonte, Schema),
    Transformacao = fxTrnAplicar(Preparacao, Schema),
    Qualidade = fxQaValidar(Transformacao, Schema),
    Validacao = fxQaFiltrarPorStatus(Qualidade, "OK"),
    Normalizacao = fxNrmAplicar(Validacao, Schema)
in
    Normalizacao;

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

shared fxTratamentoAddPrefix = (valor as any, optional parametros as nullable any) as any =>

if valor = null then

    null

else

    let

        Prefixo =
            Text.From(parametros{0}),

        Texto =
            Text.From(valor)

    in

        if Text.StartsWith(Texto, Prefixo) then

            Texto

        else

            Prefixo & Texto;

shared fxTratamentoAddSuffix = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else
        let
            Sufixo = Text.From(parametros{0}),
            Texto  = Text.From(valor)
        in
            if Text.EndsWith(Texto, Sufixo) then Texto
            else Texto & Sufixo;

shared fxTratamentoPadLeft = (valor as any, optional parametros as nullable any) as any =>
    // {tamanho, caractere, prefixo}

    if valor = null then

        null

    else

        let

            Texto =
                Text.From(valor),

            Tamanho =
                Number.From(parametros{0}),

            Caractere =
                Text.From(parametros{1}),

            Prefixo =
                if List.Count(parametros) >= 3 then
                    Text.From(parametros{2})
                else
                    null,

            Resultado =

                if Prefixo <> null and Text.StartsWith(Texto, Prefixo) then

                    Prefixo &
                    Text.PadStart(
                        Text.AfterDelimiter(Texto, Prefixo),
                        Tamanho,
                        Caractere
                    )

                else

                    Text.PadStart(
                        Texto,
                        Tamanho,
                        Caractere
                    )

        in

            Resultado;

shared fxTratamentoPadRight = (valor as any, optional parametros as nullable any) as any =>
// {tamanho, caractere, sufixo}

if valor = null then

    null

else

    let

        Texto =
            Text.From(valor),

        Tamanho =
            Number.From(parametros{0}),

        Caractere =
            Text.From(parametros{1}),

        Sufixo =
            if List.Count(parametros) >= 3 then
                Text.From(parametros{2})
            else
                null,

        Resultado =

            if
                Sufixo <> null and
                Text.EndsWith(Texto, Sufixo)
            then

                Text.PadEnd(
                    Text.BeforeDelimiter(
                        Texto,
                        Sufixo,
                        {0, RelativePosition.FromEnd}
                    ),
                    Tamanho,
                    Caractere
                ) &
                Sufixo

            else

                Text.PadEnd(
                    Texto,
                    Tamanho,
                    Caractere
                )

    in

        Resultado;

shared fxTratamentoRemoveChars = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else
        Text.Remove(Text.From(valor), parametros);

shared fxTratamentoKeepChars = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else
        Text.Select(Text.From(valor), parametros);

// Mapa de acentos como constante compartilhada — avaliado uma vez por refresh, não por linha.
shared srcMapaAcentos = List.Buffer({
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
});
shared fxTratamentoRemoveAccents = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then
        null
    else
        List.Accumulate(
            srcMapaAcentos,
            Text.From(valor),
            (Estado, Item) => Text.Replace(Estado, Item{0}, Item{1})
        );

shared fxTratamentoRemovePunctuation = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null else
        Text.Remove(Text.From(valor), ".,;:!?()[]{}<>/\|-_""'");

shared fxTratamentoKeepText = (valor as any, optional parametros as nullable any) as any =>
    if valor = null then null
    else Text.Select(Text.From(valor), srcKeepTextChars);

shared fxOperadoresPadrao = // Retorna {Tratamentos, Validações} padrão para o tipo de dado informado.
(Tipo as type) as record =>
let
    Chave = fxTipoParaTexto(Tipo)
in
    [
        Tratamentos = Record.FieldOrDefault(cfgOperadoresTratamentoPadraoMap, Chave, {}),
        Validações  = Record.FieldOrDefault(cfgOperadoresValidacaoPadraoMap,  Chave, {})
    ];

shared fxTratamentoNumber = 
// Converte texto numérico para number respeitando separadores decimal e de milhar da cultura.
(valor as any, optional parametros as nullable any) as any =>
    fxConvertTextToNumber(valor, cfgCultura);

shared parCulturaBootstrap = 
// Lê Cultura_Padrao diretamente da fonte para quebrar o ciclo cfgCultura → cfgParametros → stgParametrosExcel.
let
    Linha =
        Table.SelectRows(
            try srcParametrosExcel otherwise #table({"Parâmetro", "Valor"}, {}),
            each [Parâmetro] = "Cultura_Padrao"
        ),
    Valor =
        if Table.IsEmpty(Linha) then null
        else Linha{0}[Valor]
in
    if Valor = null or Text.Trim(Text.From(Valor)) = ""
    then "pt-BR"
    else Text.Trim(Text.From(Valor));

shared fxConvertTextToNumber = (valor as any, cultura as record) as nullable number =>
let
    Resultado =
        if valor = null then

            null

        else if Value.Is(valor, type number) then

            valor

        else if Value.Is(valor, type text) then

            let
                Texto =
                    Text.Trim(valor),

                Negativo =
                    cultura[NegativoParenteses]
                    and Text.StartsWith(Texto, "(")
                    and Text.EndsWith(Texto, ")"),

                TextoSemParenteses =
                    if Negativo then
                        Text.Middle(Texto, 1, Text.Length(Texto) - 2)
                    else
                        Texto,

                TextoLimpo =
                    Text.Select(
                        TextoSemParenteses,
                        cultura[CaracteresPermitidosNumero]
                    ),

                Numero =
                    if TextoLimpo = "" then
                        error "Valor numérico vazio."
                    else
                        try
                            Number.FromText(
                                TextoLimpo,
                                cultura[Cultura]
                            )
                        otherwise
                            error "Valor numérico inválido."
            in
                if Negativo then
                    -Numero
                else
                    Numero

        else

            error "Tipo de entrada inválido. Esperado texto ou número."
in
    Resultado;

shared tstClientes_SemFramework = let
    Fonte = srcDados,

    // Tratamentos
    Tratamentos =
        Table.TransformColumns(
            Fonte,
            {
                {
                    "CPF",
                    each
                        let
                            Digitos = Text.Select(Text.From(_), {"0".."9"})
                        in
                            Digitos,
                    type text
                },
                {
                    "Nome",
                    each
                        Text.Proper(
                            Text.Select(
                                Text.From(_),
                                {"A".."Z","a".."z","À".."Ö","Ø".."ö","ø".."ÿ"," "}
                            )
                        ),
                    type text
                },
                {
                    "RG",
                    each Text.Select(Text.From(_), {"0".."9"}),
                    type text
                },
                {
                    "DataNascimento",
                    each try Date.From(_) otherwise null,
                    type date
                },
                {
                    "Sexo",
                    each Text.Proper(Text.From(_)),
                    type text
                },
                {
                    "EstadoCivil",
                    each Text.Proper(Text.From(_)),
                    type text
                },
                {
                    "Email",
                    each Text.Lower(Text.Trim(Text.From(_))),
                    type text
                },
                {
                    "Telefone",
                    each Text.Select(Text.From(_), {"0".."9"}),
                    type text
                },
                {
                    "Celular",
                    each Text.Select(Text.From(_), {"0".."9"}),
                    type text
                },
                {
                    "CEP",
                    each Text.Select(Text.From(_), {"0".."9"}),
                    type text
                },
                {
                    "Logradouro",
                    each Text.Proper(Text.From(_)),
                    type text
                },
                {
                    "Numero",
                    each try Int64.From(_) otherwise null,
                    Int64.Type
                },
                {
                    "Complemento",
                    each Text.Proper(Text.From(_)),
                    type text
                },
                {
                    "Bairro",
                    each
                        Text.Proper(
                            Text.Select(
                                Text.From(_),
                                {"A".."Z","a".."z","À".."Ö","Ø".."ö","ø".."ÿ"," "}
                            )
                        ),
                    type text
                },
                {
                    "Cidade",
                    each
                        Text.Proper(
                            Text.Select(
                                Text.From(_),
                                {"A".."Z","a".."z","À".."Ö","Ø".."ö","ø".."ÿ"," "}
                            )
                        ),
                    type text
                },
                {
                    "UF",
                    each Text.Upper(Text.From(_)),
                    type text
                },
                {
                    "Profissao",
                    each
                        Text.Combine(
                            List.Select(
                                Text.Split(
                                    Text.Proper(
                                        Text.Trim(
                                            Text.Select(
                                                Text.From(_),
                                                {"A".."Z","a".."z","À".."Ö","Ø".."ö","ø".."ÿ"," "}
                                            )
                                        )
                                    ),
                                    " "
                                ),
                                each _ <> ""
                            ),
                            " "
                        ),
                    type text
                },
                {
                    "RendaMensal",
                    each try Number.From(_) otherwise null,
                    type number
                },
                {
                    "DataCadastro",
                    each try Date.From(_) otherwise null,
                    type date
                },
                {
                    "ClienteAtivo",
                    each if (_) = "Sim" then true else false,
                    type logical
                }
            }
        ),

    // Validações
    ValidarSexo =
        Table.AddColumn(
            Tratamentos,
            "ValidoSexo",
            each
                List.Contains(
                    {"Masculino","Feminino","Não Informado"},
                    [Sexo]
                ),
            type logical
        ),

    ValidarEstadoCivil =
        Table.AddColumn(
            ValidarSexo,
            "ValidoEstadoCivil",
            each
                List.Contains(
                    {
                        "Casado",
                        "Divorciado",
                        "Não Informado",
                        "Solteiro",
                        "União Estável",
                        "Viúvo"
                    },
                    [EstadoCivil]
                ),
            type logical
        ),

    ValidarEmail =
        Table.AddColumn(
            ValidarEstadoCivil,
            "ValidoEmail",
            each
                let
                    Email = [Email]
                in
                    Email <> null
                    and Text.Contains(Email, "@")
                    and Text.Contains(Text.AfterDelimiter(Email, "@"), "."),
            type logical
        ),

    Erros =
        Table.AddColumn(
            ValidarEmail,
            "Ocorrencias",
            each
                List.RemoveNulls({
                    if not List.Contains({"Masculino","Feminino","Não Informado"}, [Sexo]) then "Sexo inválido" else null,

                    if not List.Contains({"Casado","Divorciado","Não Informado","Solteiro","União Estável","Viúvo"}, [EstadoCivil]) then "Estado Civil inválido" else null,

                    if [Email] = null
                        or not Text.Contains([Email], "@")
                        or not Text.Contains(Text.AfterDelimiter([Email], "@"), ".")
                    then "E-mail inválido"
                    else null
                }),
            type list
        ),

    RemoverRegistrosInvalidos =
        Table.SelectRows(
            Erros,
            each List.IsEmpty([Ocorrencias])
        ),
    ColunasRemovidas = Table.RemoveColumns(RemoverRegistrosInvalidos,{"ValidoSexo", "ValidoEstadoCivil", "ValidoEmail", "Ocorrencias"})
in
    ColunasRemovidas;

shared fxTratamentoRemovePrefix = (valor as any, optional parametros as nullable any) as any =>
let
    Texto =
        if valor = null then
            null
        else
            Text.From(valor),

    Prefixos =
        if parametros = null then
            {}
        else if Value.Is(parametros, type list) then
            parametros
        else
            {Text.From(parametros)},

    PrefixosOrdenados =
        List.Sort(
            Prefixos,
            (x, y) => Value.Compare(Text.Length(y), Text.Length(x))
        ),

    Resultado =
        if Texto = null then
            null
        else
            List.Accumulate(
                PrefixosOrdenados,
                Texto,
                (estado, prefixo) =>
                    if Text.StartsWith(estado, prefixo) then
                        Text.TrimStart(
                            Text.Range(
                                estado,
                                Text.Length(prefixo)
                            )
                        )
                    else
                        estado
            )
in
    Resultado;

shared fxTratamentoRemoveSuffix = (valor as any, optional parametros as nullable any) as any =>
let
    Texto =
        if valor = null then
            null
        else
            Text.From(valor),

    Sufixos =
        if parametros = null then
            {}
        else if Value.Is(parametros, type list) then
            parametros
        else
            {Text.From(parametros)},

    SufixosOrdenados =
        List.Sort(
            Sufixos,
            (x, y) => Value.Compare(Text.Length(y), Text.Length(x))
        ),

    Resultado =
        if Texto = null then
            null
        else
            List.Accumulate(
                SufixosOrdenados,
                Texto,
                (estado, sufixo) =>
                    if Text.EndsWith(estado, sufixo) then
                        Text.TrimEnd(
                            Text.Start(
                                estado,
                                Text.Length(estado) - Text.Length(sufixo)
                            )
                        )
                    else
                        estado
            )
in
    Resultado;

shared srcKeepTextChars = List.Buffer(
    List.Distinct(
        Text.ToList("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 ")
        & List.Transform(srcMapaAcentos, each _{0})
    )
);

shared fxValidacaoPhone = (
    valor as any,
    optional parametros as nullable list,
    optional contexto as nullable record
)
as record =>

let
    Digitos =
        if valor = null then
            null
        else
            Text.Select(
                Text.From(valor),
                {"0".."9"}
            ),

    Numero =
        if Digitos = null then
            null
        else if List.Contains({12, 13}, Text.Length(Digitos)) and Text.StartsWith(Digitos, "55") then
            Text.Range(Digitos, 2)
        else
            Digitos,

    Tamanho =
        if Numero = null then
            null
        else
            Text.Length(Numero),

    PrimeiroDigito =
        if Numero = null then
            null
        else
            Text.At(Numero, 2),

    Valido =
        if valor = null then
            true
        else if Tamanho = 10 then
            List.Contains({"2", "3", "4", "5"}, PrimeiroDigito)
        else if Tamanho = 11 then
            PrimeiroDigito = "9"
        else
            false,

    Ocorrencias =
        if Valido then
            null
        else
            {
                fxSchemaOcorrencia(
                    "PHONEVAL",
                    contexto,
                    valor,
                    valor,
                    "Telefone inválido.",
                    []
                )
            }

in
    [
        Valor = valor,
        Ocorrencias = Ocorrencias
    ];

shared fxFeriadoDataReferencia = (Ano as number, Referência as text) as nullable date =>
let
    Ref = Text.Upper(Text.Trim(Referência)),

    Páscoa =
        let
            A = Number.Mod(Ano, 19),
            B = Number.IntegerDivide(Ano, 100),
            C = Number.Mod(Ano, 100),
            D = Number.IntegerDivide(B, 4),
            E = Number.Mod(B, 4),
            F = Number.IntegerDivide(B + 8, 25),
            G = Number.IntegerDivide(B - F + 1, 3),
            H = Number.Mod(19 * A + B - D - G + 15, 30),
            I = Number.IntegerDivide(C, 4),
            K = Number.Mod(C, 4),
            L = Number.Mod(32 + 2 * E + 2 * I - H - K, 7),
            M = Number.IntegerDivide(A + 11 * H + 22 * L, 451),
            Mes = Number.IntegerDivide(H + L - 7 * M + 114, 31),
            Dia = Number.Mod(H + L - 7 * M + 114, 31) + 1
        in
            #date(Ano, Mes, Dia),

    SegundoDomingoMaio =
        let
            Inicio = #date(Ano, 5, 1),
            PrimeiroDomingo = Date.AddDays(Inicio, Number.Mod(7 - Date.DayOfWeek(Inicio, Day.Sunday), 7))
        in
            Date.AddDays(PrimeiroDomingo, 7),

    SegundoDomingoAgosto =
        let
            Inicio = #date(Ano, 8, 1),
            PrimeiroDomingo = Date.AddDays(Inicio, Number.Mod(7 - Date.DayOfWeek(Inicio, Day.Sunday), 7))
        in
            Date.AddDays(PrimeiroDomingo, 7),

    Data =
        if Ref = "PASCOA" then
            Páscoa
        else if Ref = "SEGUNDO_DOMINGO_MAIO" then
            SegundoDomingoMaio
        else if Ref = "SEGUNDO_DOMINGO_AGOSTO" then
            SegundoDomingoAgosto
        else
            null
in
    Data;

shared fxFeriadoGerar = (Ano as number, optional Estado as nullable text, optional Municipio as nullable text) as table =>
let
    Fonte = cfgFeriados[Todos],

    FiltrarVigencia =
        Table.SelectRows(
            Fonte,
            each
                [VigênciaInicial] <= #date(Ano, 12, 31) and
                (
                    [VigênciaFinal] = null or
                    [VigênciaFinal] >= #date(Ano, 1, 1)
                )
        ),

    FiltrarEstado =
        if Estado = null then
            FiltrarVigencia
        else
            Table.SelectRows(
                FiltrarVigencia,
                each [Estado] = null or [Estado] = Estado
            ),

    FiltrarMunicipio =
        if Municipio = null then
            FiltrarEstado
        else
            Table.SelectRows(
                FiltrarEstado,
                each [Município] = null or [Município] = Municipio
            ),

    AdicionarData =
        Table.AddColumn(
            FiltrarMunicipio,
            "Data",
            each
                if [Regularidade] = "FIXO" then
                    #date(Ano, [Mês], [Dia])
                else
                    Date.AddDays(
                        fxFeriadoDataReferencia(Ano, [Referência]),
                        [Offset]
                    ),
            type date
        ),

    Selecionar =
        Table.SelectColumns(
            AdicionarData,
            {
                "Data",
                "Nome",
                "Código",
                "Abrangência",
                "PontoFacultativo",
                "Expediente",
                "Estado",
                "Município"
            }
        ),

    Ordenar =
        Table.Sort(
            Selecionar,
            {
                {"Data", Order.Ascending},
                {"Nome", Order.Ascending}
            }
        )

in
    Ordenar;

shared fxFeriado = (
    Data as date
)
as record =>

let
    Ano =
        Text.From(
            Date.Year(Data)
        ),

    IndiceAno =
        Record.FieldOrDefault(
            cfgFeriadosIndices,
            Ano,
            []
        ),

    Resultado =
        Record.FieldOrDefault(
            IndiceAno,
            Date.ToText(Data, "yyyy-MM-dd"),
            [
                Existe = false,
                Nome = null,
                Código = null,
                Abrangência = null,
                PontoFacultativo = false,
                Expediente = null,
                Estado = null,
                Município = null
            ]
        )

in
    Resultado;

shared stgDados = let
    Fonte = srcDados,
    Preparada = fxStgAplicar(Fonte, "nrmDados"),
    Resultado = Preparada
in
    Resultado;

shared diagPipelineEstrutura = let
    Fonte =
        stgTabelasPowerQuery,

    Consultas =
        Table.SelectRows(
            Fonte,
            each [Origem] = "PowerQuery"
        ),

    Entidade =
        Table.AddColumn(
            Consultas,
            "Entidade",
            each Text.AfterDelimiter([Nome], [Prefix]),
            type text
        ),

    Selecionar =
        Table.SelectColumns(
            Entidade,
            {"Entidade", "Categoria", "Nome"}
        ),

    Pivot =
        Table.Pivot(
            Selecionar,
            List.Distinct(Selecionar[Categoria]),
            "Categoria",
            "Nome"
        ),

    Categorias =
        Record.ToTable(cfgCategoriasPowerQuery),

    CategoriasOrdenadas =
        Table.Sort(
            Table.ExpandRecordColumn(
                Categorias,
                "Value",
                {"Ordem", "Categoria"}
            ),
            {{"Ordem", Order.Ascending}}
        ),

    Ordem =
        Table.ReorderColumns(
            Pivot,
            List.Combine(
                {
                    {"Entidade"},
                    CategoriasOrdenadas[Categoria]
                }
            ),
            MissingField.Ignore
        ),

    Resultado =
        Table.Sort(
            Ordem,
            {{"Entidade", Order.Ascending}}
        ),

    Filtrar =
        Table.SelectRows(
            Resultado,
            each not (
                Record.FieldOrDefault(_, "Normalização", null) = null and
                Record.FieldOrDefault(_, "Dimensão", null) = null and
                Record.FieldOrDefault(_, "Fato", null) = null
            )
        )

in
    Filtrar;

shared fxResolverCaminho = (Caminho as text) as record =>

let
    Valor =
        Text.Trim(Caminho),

    Normalizado =
        Text.Replace(
            Valor,
            "/",
            "\"
        ),

    CaminhoSemBarraFinal =
        if Text.EndsWith(Normalizado, "\") then
            Text.Start(
                Normalizado,
                Text.Length(Normalizado) - 1
            )
        else
            Normalizado,

    UltimoSeparador =
        Text.PositionOf(
            CaminhoSemBarraFinal,
            "\",
            Occurrence.Last
        ),

    NomeUltimoElemento =
        if UltimoSeparador >= 0 then
            Text.Range(
                CaminhoSemBarraFinal,
                UltimoSeparador + 1
            )
        else
            CaminhoSemBarraFinal,

    Extensao =
        if Text.Contains(NomeUltimoElemento, ".") then
            Text.AfterDelimiter(
                NomeUltimoElemento,
                ".",
                {0, RelativePosition.FromEnd}
            )
        else
            null,

    EhArquivo =
        Extensao <> null,

    Pasta =
        if EhArquivo then
            Text.Start(
                CaminhoSemBarraFinal,
                UltimoSeparador + 1
            )
        else
            CaminhoSemBarraFinal & "\",

    Arquivo =
        if EhArquivo then
            NomeUltimoElemento
        else
            null

in
    [
        Pasta = Pasta,
        Arquivo = Arquivo
    ];

shared fxPipelineDiagnostico = (
    Fonte as table,
    STG as table,
    TRN as table,
    QA as table,
    NRM as table
)
as table =>

let
    NormalizarOcorrencias =
        (Valor as any) as list =>
            if Valor = null then
                {}
            else if Value.Is(Valor, type list) then
                Valor
            else if Value.Is(Valor, type record) then
                {Valor}
            else
                {},

    ContarOcorrencias =
        (Tabela as table, Coluna as text) as record =>
            if not Table.HasColumns(Tabela, Coluna) then
                [
                    RegistrosComOcorrencias = 0,
                    TotalOcorrencias = 0
                ]
            else
                let
                    Valores =
                        Table.Column(
                            Tabela,
                            Coluna
                        ),

                    Listas =
                        List.Transform(
                            Valores,
                            each NormalizarOcorrencias(_)
                        ),

                    RegistrosComOcorrencias =
                        List.Count(
                            List.Select(
                                Listas,
                                each not List.IsEmpty(_)
                            )
                        ),

                    TotalOcorrencias =
                        List.Sum(
                            List.Transform(
                                Listas,
                                each List.Count(_)
                            )
                        )
                in
                    [
                        RegistrosComOcorrencias =
                            RegistrosComOcorrencias,
                        TotalOcorrencias =
                            TotalOcorrencias
                    ],

    ContarStatus =
        (Tabela as table) as record =>
            if not Table.HasColumns(
                Tabela,
                "_QA_Status"
            ) then
                [
                    OK = null,
                    AVISO = null,
                    ERRO = null
                ]
            else
                let
                    Valores =
                        Table.Column(
                            Tabela,
                            "_QA_Status"
                        )
                in
                    [
                        OK =
                            List.Count(
                                List.Select(
                                    Valores,
                                    each _ = "OK"
                                )
                            ),
                        AVISO =
                            List.Count(
                                List.Select(
                                    Valores,
                                    each _ = "AVISO"
                                )
                            ),
                        ERRO =
                            List.Count(
                                List.Select(
                                    Valores,
                                    each _ = "ERRO"
                                )
                            )
                    ],

    CriarResumo =
        (
            Etapa as text,
            Entrada as table,
            Saida as table,
            ColunaOcorrencias as nullable text
        ) as record =>
            let
                TotalEntrada =
                    Table.RowCount(
                        Entrada
                    ),

                TotalSaida =
                    Table.RowCount(
                        Saida
                    ),

                Ocorrencias =
                    if ColunaOcorrencias = null then
                        [
                            RegistrosComOcorrencias = 0,
                            TotalOcorrencias = 0
                        ]
                    else
                        ContarOcorrencias(
                            Saida,
                            ColunaOcorrencias
                        ),

                Status =
                    ContarStatus(
                        Saida
                    )
            in
                [
                    Etapa = Etapa,
                    RegistrosEntrada = TotalEntrada,
                    RegistrosSaida = TotalSaida,
                    RegistrosRemovidos =
                        TotalEntrada - TotalSaida,
                    Retencao =
                        if TotalEntrada = 0 then
                            null
                        else
                            TotalSaida / TotalEntrada,
                    RegistrosComOcorrencias =
                        Ocorrencias[RegistrosComOcorrencias],
                    TotalOcorrencias =
                        Ocorrencias[TotalOcorrencias],
                    OK = Status[OK],
                    AVISO = Status[AVISO],
                    ERRO = Status[ERRO]
                ],

    ResumoFonte =
        [
            Etapa = "SRC",
            RegistrosEntrada = null,
            RegistrosSaida =
                Table.RowCount(Fonte),
            RegistrosRemovidos = null,
            Retencao = null,
            RegistrosComOcorrencias = 0,
            TotalOcorrencias = 0,
            OK = null,
            AVISO = null,
            ERRO = null
        ],

    ResumoSTG =
        CriarResumo(
            "STG",
            Fonte,
            STG,
            "_STG_Ocorrencias"
        ),

    ResumoTRN =
        CriarResumo(
            "TRN",
            STG,
            TRN,
            null
        ),

    ResumoQA =
        CriarResumo(
            "QA",
            TRN,
            QA,
            "_QA_Ocorrencias"
        ),

    ResumoNRM =
        CriarResumo(
            "NRM",
            QA,
            NRM,
            null
        ),

    Resultado =
        Table.FromRecords(
            {
                ResumoFonte,
                ResumoSTG,
                ResumoTRN,
                ResumoQA,
                ResumoNRM
            }
        ),

    Tipos =
        Table.TransformColumnTypes(
            Resultado,
            {
                {"Etapa", type text},
                {"RegistrosEntrada", Int64.Type},
                {"RegistrosSaida", Int64.Type},
                {"RegistrosRemovidos", Int64.Type},
                {"Retencao", Percentage.Type},
                {"RegistrosComOcorrencias", Int64.Type},
                {"TotalOcorrencias", Int64.Type},
                {"OK", Int64.Type},
                {"AVISO", Int64.Type},
                {"ERRO", Int64.Type}
            }
        )

in
    Tipos;

shared fxDados = (
    optional Config as nullable record
)
as table =>

let
    Fonte =
        fxOrigem(
            Config
        ),

    FonteDados =
        Text.Upper(
            fxParametro(
                "Dados_Fonte",
                null,
                null,
                Config
            )
        ),

    Tabela =
        if
            Value.Is(
                Fonte,
                type table
            )
            and Table.HasColumns(
                Fonte,
                "Dados"
            )
            and not Table.IsEmpty(
                Fonte
            )
            and Value.Is(
                Fonte{0}[Dados],
                type table
            )
        then

            Table.Combine(
                Fonte[Dados]
            )

        else if
            FonteDados = "ARQUIVOS"
            and Value.Is(
                Fonte,
                type record
            )
            and Record.HasFields(
                Fonte,
                "Arquivos"
            )
            and Record.HasFields(
                Fonte,
                "Arquivo"
            )
        then

            let
                Origem =
                    Fonte,

                FormatoArquivo =
                    Text.Upper(
                        Text.From(
                            fxParametro(
                                "Arquivo_Formato",
                                "",
                                null,
                                Config
                            )
                        )
                    ),

                DelimitadorArquivo =
                    fxParametro(
                        "Arquivo_Delimitador",
                        ";",
                        null,
                        Config
                    ),

                CodificadorArquivo =
                    fxParametro(
                        "Arquivo_Codificador",
                        65001,
                        null,
                        Config
                    ),

                QuoteStyleArquivo =
                    fxParametro(
                        "Arquivo_QuoteStyle",
                        QuoteStyle.Csv,
                        null,
                        Config
                    ),

                PromoverCabecalhos =
                    fxParametro(
                        "Arquivo_Promover_Cabecalhos",
                        true,
                        null,
                        Config
                    ),

                ArquivosFiltrados =
                    fxFiltrarArquivos(
                        Origem[Arquivos],
                        FormatoArquivo,
                        Origem[Arquivo]
                    ),

                Dados =
                    Table.AddColumn(
                        ArquivosFiltrados,
                        "Dados",
                        each
                            fxLeitorArquivo(
                                [Content],
                                FormatoArquivo,
                                DelimitadorArquivo,
                                CodificadorArquivo,
                                QuoteStyleArquivo,
                                PromoverCabecalhos
                            ),
                        type any
                    )

            in
                if
                    Table.IsEmpty(
                        Dados
                    )
                then
                    #table(
                        {},
                        {}
                    )
                else
                    Table.Combine(
                        Dados[Dados]
                    )

        else

            fxOrigemComoTabela(
                Fonte
            ),

    LinhasValidas =
        Table.SelectRows(
            Tabela,
            each
                List.NonNullCount(
                    Record.FieldValues(_)
                ) > 0
        )

in
    LinhasValidas;

shared fxObjetoIdentificarPrefixo = (objeto as text) as nullable text =>
let
    ObjetoLower = Text.Lower(objeto),
    Prefixos = cfgPrefixosPowerQuery,
    // Ordena por tamanho decrescente para priorizar prefixos mais longos
    PrefixosOrdenados = List.Sort(Prefixos, (a,b) => Value.Compare(Text.Length(b), Text.Length(a))),
    PrefixoEncontrado = List.First(
        List.Select(
            PrefixosOrdenados,
            each Text.StartsWith(ObjetoLower, _)
        ),
        null
    )
in
    PrefixoEncontrado;