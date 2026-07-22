// Power Query from: Excel Data Framework DBB.xlsx
// Pathname: c:\Users\daniel-bighelini\OneDrive\Documentos\Planilhas\Excel Data Framework DBB\Excel Data Framework DBB.xlsx
// Extracted: 2026-07-22T14:45:41.040Z

section Section1;

shared srcSections = let
    Fonte =
        Record.RemoveFields(
            #sections[Section1],
            {"srcSections", "diagSections", "diagConsultasPQ", "diagTabelasExcel"}
        )
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

shared srcParametrosPowerQuery = let
    Fonte = srcSections,

    // EXCLUI CONSULTAS DE INFRAESTRUTURA que dependem de srcParametrosPowerQuery
    // eliminando a referência cíclica: stgParametros -> srcParametrosPowerQuery -> stgParametros
    NomesCandidatos = List.Select(
        Record.FieldNames(Fonte),
        (n) => not (
            Text.StartsWith(n, "diag") or 
            Text.StartsWith(n, "stg") or 
            Text.StartsWith(n, "cfg") or 
            Text.StartsWith(n, "src") or 
            Text.StartsWith(n, "fx") or 
            Text.StartsWith(n, "dim") or 
            Text.StartsWith(n, "fato") or 
            Text.StartsWith(n, "nrm")
        )
    ),

    ConsultasCandidatas = Record.SelectFields(Fonte, NomesCandidatos, MissingField.Ignore),

    NomesParametros = List.Select(
        Record.FieldNames(ConsultasCandidatas),
        (Nome) =>
            let
                Valor = Record.Field(ConsultasCandidatas, Nome),
                Metadata = try Value.Metadata(Valor) otherwise null
            in
                Record.FieldOrDefault(Metadata, "IsParameterQuery", false) = true
    ),

    Parametros = Record.SelectFields(ConsultasCandidatas, NomesParametros, MissingField.Ignore)
in
    Parametros;

shared srcParametrosTipos = let
    Fonte = srcWorkbook{[Name=parTabelaParametrosTipos]}[Content]
in
    Fonte;

shared srcParametrosBooleanos = let
    Fonte = srcWorkbook{[Name=parTabelaParametrosBooleanos]}[Content]
in
    Fonte;

shared srcParametrosFormatosArquivos = let
    Fonte = srcWorkbook{[Name=parTabelaParametrosFormatosArquivos]}[Content]
in
    Fonte;

shared srcParametrosRESTHeaders = let
    Fonte = srcWorkbook{[Name=parTabelaParametrosRESTHeaders]}[Content]
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

shared srcExcelClientes = let
    Fonte = srcWorkbook{[Name=parTabelaExcelClientes]}[Content]
in
    Fonte;

shared srcExcelProdutos = let
    Fonte = srcWorkbook{[Name=parTabelaExcelProdutos]}[Content]
in
    Fonte;

shared srcExcelVendas = let
    Fonte = srcWorkbook{[Name=parTabelaExcelVendas]}[Content]
in
    Fonte;

shared srcExcelDadosExternos = let
    Fonte = srcWorkbook{[Name=parTabelaExcelDadosExternos]}[Content]
in
    Fonte;

shared parTabelaParametros = "tbParametros" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaParametrosBooleanos = "tbParametrosBooleanos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaParametrosTipos = "tbParametrosTipos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared cfgParametros = // Materializa a tabela consolidada em um Record para permitir
// acesso O(1) aos parâmetros através de Record.Field.
Record.FromTable(
    Table.Buffer(stgParametros)
);

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

shared stgParametros = let
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
                Record.FieldNames(srcParametrosPowerQuery),
                (Nome) =>
                    let
                        Valor =
                            Record.Field(
                                srcParametrosPowerQuery,
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

    // OTIMIZAÇÃO: Materialização do catálogo sem Table.ToRecords
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
    ParametrosCatalogo;

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
        cfgParametrosTipos,
        Nome
    ) then

        Record.Field(
            cfgParametrosTipos,
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
                                cfgParametrosTipos
                            )
                        ),
                        ", "
                    )
            ]
        );

shared cfgParametrosTipos = Record.FromTable(
    Table.Buffer(stgParametrosTipos)
)
;

shared stgParametrosTipos = let
    Fonte = srcParametrosTipos,
    LinhasEmBrancoRemovidas = Table.SelectRows(Fonte, each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))),

    Tipos =
        Table.TransformColumns(
            LinhasEmBrancoRemovidas,
            {
                {
                    "Descrição",
                    each
                        if _ = null then
                            null
                        else
                            Text.Upper(Text.Trim(_)),
                    type nullable text
                },
                {
                    "Código",
                    each
                        if _ = null then
                            null
                        else
                            Text.Upper(Text.Trim(_)),
                    type nullable text
                }
            }
        ),

    // OTIMIZAÇÃO DIRETA: Construção de tabelas projetadas via vetores de colunas
    BaseExcel  = Table.FromColumns({Tipos[Descrição], Tipos[Código]}, {"Name", "Codigo"}),
    BaseCodigo = Table.FromColumns({Tipos[Código], Tipos[Código]}, {"Name", "Codigo"}),
    TiposExpandido = Table.Combine({BaseExcel, BaseCodigo}),
    
    TiposUnicos =
        Table.Distinct(
            TiposExpandido,
            {"Name"}
        ),

    Mapeamento =
        Table.AddColumn(
            TiposUnicos,
            "Value",
            each
                if [Codigo] = "ANY" then type any
                else if [Codigo] = "TEXT" then type text
                else if [Codigo] = "INT64" then Int64.Type
                else if [Codigo] = "NUMBER" then type number
                else if [Codigo] = "LOGICAL" then type logical
                else if [Codigo] = "DATE" then type date
                else if [Codigo] = "DATETIME" then type datetime
                else if [Codigo] = "DATETIMEZONE" then type datetimezone
                else if [Codigo] = "DURATION" then type duration
                else if [Codigo] = "TIME" then type time
                else if [Codigo] = "LIST" then type list
                else
                    error Error.Record(
                        "Tipo do Power Query inválido",
                        "O tipo informado na coluna 'PQ' não é suportado.",
                        [Codigo = [Codigo]]
                    ),
            type nullable type
        ),

    Resultado =
        Table.SelectColumns(
            Mapeamento,
            {
                "Name",
                "Value"
            }
        )
in
    Resultado;

shared stgParametrosBooleanos = let
    Fonte = srcParametrosBooleanos,

    Renomeado = Table.RenameColumns(
        Fonte,
        {
            {"Descrição", "Name"},
            {"Código", "Value"}
        }
    ),

    VaziosRemovidos = Table.SelectRows(
        Renomeado,
        each [Name] <> null
            and Text.Trim(Text.From([Name])) <> ""
    ),
    DuplicatasRemovidas = Table.Distinct(
                            VaziosRemovidos,
                            {"Name"}
                          ),

    Normalizado =
        Table.TransformColumns(
            DuplicatasRemovidas,
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
                    each Logical.From(_),
                    type logical
                }
            }
        )
in
    Normalizado;

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

shared stgParametrosRESTHeaders = let
    Fonte =
        srcParametrosRESTHeaders,

    ColunasRenomeadas =
        Table.RenameColumns(
            Fonte,
            {
                {"Header", "Name"},
                {"Valor", "Value"}
            }
        ),

    LinhasValidas =
        Table.SelectRows(
            ColunasRenomeadas,
            each
                [Name] <> null
                and Text.Trim(Text.From([Name])) <> ""
        ),

    Normalizado =
        Table.TransformColumns(
            LinhasValidas,
            {
                {
                    "Name",
                    each Text.Trim(Text.From(_)),
                    type text
                },
                {
                    "Value",
                    each
                        if _ = null then
                            ""
                        else
                            Text.From(_),
                    type text
                }
            }
        ),

    DuplicatasRemovidas =
        Table.Distinct(
            Normalizado,
            {"Name"}
        )

in
    DuplicatasRemovidas;

shared stgParametrosTratamentos = let
    Fonte = srcParametrosTratamentos,

    RemoverVazios = fxStgRemoverVazios(Fonte, {"Ativo"}),
    Padronizado = fxStgPadronizar(RemoverVazios),
    TratamentosAtivos = Table.SelectRows(Padronizado, each [Ativo] = true),
    Funcoes = Record.FromTable(diagSections),

    Tratamentos =
        Table.TransformColumns(
            TratamentosAtivos,
            {
                {
                    "Descrição",
                    each if _ = null then null else Text.Trim(Text.From(_)),
                    type nullable text
                },
                {
                    "Código",
                    each if _ = null then null else Text.Upper(Text.Trim(Text.From(_))),
                    type nullable text
                },
                {
                    "Função",
                    each if _ = null then null else Record.FieldOrDefault(Funcoes, Text.Trim(Text.From(_)), null),
                    type function
                }
            }
        ),

    // OTIMIZAÇÃO DIRETA: Projeção de colunas com Table.FromColumns sem Table.ToRecords
    BaseExcel = Table.FromColumns(
        {
            Tratamentos[Descrição],
            Tratamentos[Código],
            Tratamentos[Função],
            Tratamentos[Ativo]
        },
        {"Name", "Código", "Função", "Ativo"}
    ),

    BaseCodigo = Table.FromColumns(
        {
            Tratamentos[Código],
            Tratamentos[Código],
            Tratamentos[Função],
            Tratamentos[Ativo]
        },
        {"Name", "Código", "Função", "Ativo"}
    ),

    TratamentosExpandido = Table.Combine({BaseExcel, BaseCodigo}),

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

    RemoverVazios =
        fxStgRemoverVazios(
            Fonte,
            {"Ativo"}
        ),

    Padronizado =
        fxStgPadronizar(
            RemoverVazios
        ),

    ValidacoesAtivos =
        Table.SelectRows(
            Padronizado,
            each [Ativo] = true
        ),

    Funcoes =
        Record.FromTable(
            diagSections
        ),

    Validacoes =
        Table.TransformColumns(
            ValidacoesAtivos,
            {
                {
                    "Descrição",
                    each
                        if _ = null then
                            null
                        else
                            Text.Trim(
                                Text.From(_)
                            ),
                    type nullable text
                },
                {
                    "Código",
                    each
                        if _ = null then
                            null
                        else
                            Text.Upper(
                                Text.Trim(
                                    Text.From(_)
                                )
                            ),
                    type nullable text
                },
                {
                    "Função",
                    each
                        if _ = null then
                            null
                        else
                            Record.FieldOrDefault(
                                Funcoes,
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
                        if _ = null then
                            null
                        else
                            Text.Upper(
                                Text.Trim(
                                    Text.From(_)
                                )
                            ),
                    type nullable text
                }
            }
        ),

    ValidacoesCompleto =
        Table.InsertRows(
            Validacoes,
            0,
            {
                [
                    Descrição = "Obrigatório",
                    Código = "REQUIRED",
                    Função = fxSchemaValidacaoREQUIRED,
                    Severidade = "ERROR",
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
                ValidacoesCompleto[Ativo]
            },
            {
                "Name",
                "Código",
                "Função",
                "Severidade",
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
                ValidacoesCompleto[Ativo]
            },
            {
                "Name",
                "Código",
                "Função",
                "Severidade",
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

    RemoverVazios =
        fxStgRemoverVazios(
            Fonte
        ),

    Padronizado =
        fxStgPadronizar(
            RemoverVazios
        ),

    Severidades =
        Table.TransformColumns(
            Padronizado,
            {
                {
                    "Código",
                    each
                        if _ = null then
                            null
                        else
                            Text.Upper(
                                Text.Trim(
                                    Text.From(_)
                                )
                            ),
                    type nullable text
                },
                {
                    "Descrição",
                    each
                        if _ = null then
                            null
                        else
                            Text.Upper(
                                Text.Trim(
                                    Text.From(_)
                                )
                            ),
                    type nullable text
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

                            Codigos =
                                if _ = null then
                                    {}
                                else
                                    List.Distinct(
                                        List.Transform(
                                            fxLista(_),
                                            each
                                                Text.Upper(
                                                    Text.Trim(
                                                        Text.From(_)
                                                    )
                                                )
                                        )
                                    ),

                            Resultado =
                                List.Transform(
                                    Codigos,
                                    (Codigo) =>

                                        let

                                            Definicao =
                                                Record.FieldOrDefault(
                                                    cfgParametrosTratamentos,
                                                    Codigo
                                                )

                                        in

                                            if Definicao = null then

                                                error Error.Record(
                                                    "Tratamento inválido",
                                                    "Tratamento não cadastrado ou foi inativado no Schema.",
                                                    [
                                                        Código = Codigo
                                                    ]
                                                )

                                            else

                                                [
                                                    Definição = Definicao,
                                                    Regra = Codigo,
                                                    Parametros = null
                                                ]
                                )

                        in

                            if List.IsEmpty(Resultado) then
                                null
                            else
                                List.Buffer(Resultado),

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

    ColunasReordenadas;

shared cfgParametrosBooleanos = Record.FromTable(
    Table.Buffer(
        stgParametrosBooleanos
    )
)
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
        cfgParametrosBooleanos,
        Nome
    ) then

        Record.Field(
            cfgParametrosBooleanos,
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
                            Record.FieldNames(cfgParametrosBooleanos)
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

shared diagSections = let
    Fonte =
        srcSections,

    ConvertidoParaTabela = Record.ToTable(Fonte)

in
    ConvertidoParaTabela;

[ Description = "Consultas existentes do PowerQuery" ]
shared diagConsultasPQ = let
    Consultas = diagSections,

    Classificacao =
        Table.AddColumn(
            Consultas,
            "Categoria",
            each
                if Text.StartsWith([Name], "par") then "Parâmetro"
                else if Text.StartsWith([Name], "cfg") then "Configuração"
                else if Text.StartsWith([Name], "fx") then "Função"
                else if Text.StartsWith([Name], "src") then "Origem"
                else if Text.StartsWith([Name], "stg") then "Stage"
                else if Text.StartsWith([Name], "nrm") then "Normalização"
                else if Text.StartsWith([Name], "dim") then "Dimensão"
                else if Text.StartsWith([Name], "fato") then "Fato"
                else if Text.StartsWith([Name], "qa") then "Qualidade"
                else if Text.StartsWith([Name], "diag") then "Diagnóstico"
                else if Text.StartsWith([Name], "tst") then "Teste"
                else if Text.StartsWith([Name], "bench") then "Benchmark"
                else if Text.StartsWith([Name], "tmp") then "Temporário"
                else if Text.StartsWith([Name], "z") then "Arquivado"
                else "Não Classificado",
            type text
        ),

    Resultado =
        Table.SelectColumns(
            Classificacao,
            {
                "Name",
                "Categoria",
                "Value"
            }
        ),

    ColunasRenomeadas =
        Table.RenameColumns(
            Resultado,
            {
                {"Name", "Consulta"},
                {"Value", "Valor"}
            }
        ),

    ConsultasClassificadas =
        Table.Sort(
            ColunasRenomeadas,
            {
                {"Categoria", Order.Ascending},
                {"Consulta", Order.Ascending}
            }
        )
in
    ConsultasClassificadas;

shared diagTabelasExcel = let
    Tabelas =
        Table.SelectColumns(
            srcWorkbook,
            {"Name", "Content"}
        ),

    LinhasFiltradas =
        Table.SelectRows(
            Tabelas,
            each
                not Text.StartsWith([Name], "diag")
                and not Text.Contains([Name], "!")
        ),

    AdicionarColunas =
        Table.AddColumn(
            LinhasFiltradas,
            "Colunas",
            each Text.Combine(Table.ColumnNames([Content]), ";"),
            type text
        ),

    Renomeadas =
        Table.RenameColumns(
            Table.RemoveColumns(AdicionarColunas, {"Content"}),
            {
                {"Name", "Tabela"}
            }
        )
in
    Renomeadas;

shared stgExcelClientes = let
    Fonte =
        srcExcelClientes,

    Stage =
        fxStgAplicar(
            Fonte,
            parTabelaExcelClientes
        ),

    Resultado =
        Table.Distinct(
            Stage
        )
in
    Resultado;

shared stgExcelProdutos = let
    Fonte =
        srcExcelProdutos,

    Stage =
        fxStgAplicar(
            Fonte,
            parTabelaExcelProdutos
        ),

    Resultado =
        Table.Distinct(
            Stage
        )
in
    Resultado;

shared stgExcelVendas = let
    Fonte =
        srcExcelVendas,

    Preparado =
        fxStgAplicar(
            Fonte,
            parTabelaExcelVendas
        ),

    RegistrosUnicos =
        Table.Distinct(
            Preparado
        )
in
    RegistrosUnicos
;

shared nrmExcelClientes = let

    Fonte =
        stgExcelClientes,

    RegistrosValidos =
        fxSchemaDescartarRegistrosBloqueantes(
            Fonte
        )

in

    RegistrosValidos;

shared dimExcelClientes = let
    // Obtém a entidade normalizada.
    Fonte =
        nrmExcelClientes,

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

shared fatoExcelVendas = let
    // Obtém os dados normalizados.
    Fonte =
        nrmExcelVendas,
    // Resolve a chave da dimensão Cliente.
    Cliente =
        Table.NestedJoin(
            Fonte,
            {"CPF"},
            dimExcelClientes,
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
            dimExcelProdutos,
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
        cfgParametrosRESTHeaders
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

shared cfgParametrosRESTHeaders = Record.FromTable(
    Table.Buffer(
        stgParametrosRESTHeaders
    )
);

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

shared cfgParametrosTratamentos = let
    Fonte = Table.Buffer(stgParametrosTratamentos),

    Nomes = Table.Column(Fonte, "Name"),
    Valores = Table.AddColumn(
        Fonte,
        "RecordValue",
        each [
            Código = [Código],
            Função = [Função],
            Ativo = [Ativo]
        ],
        type record
    )[RecordValue],

    Resultado = Record.FromList(Valores, Nomes)
in
    Resultado;

shared cfgParametrosValidacoes = let
    Fonte = Table.Buffer(stgParametrosValidacoes),

    Nomes = Table.Column(Fonte, "Name"),
    Valores = Table.AddColumn(
        Fonte,
        "RecordValue",
        each [
            Código = [Código],
            Função = [Função],
            Severidade = [Severidade],
            Ativo = [Ativo]
        ],
        type record
    )[RecordValue],

    Resultado = Record.FromList(Valores, Nomes)
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

shared cfgSchema = let
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
    )
in
    Schema;

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

shared parTabelaParametrosRESTHeaders = "tbParametrosRESTHeaders" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaParametrosTratamentos = "tbParametrosTratamentos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaParametrosValidacoes = "tbParametrosValidacoes" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaParametrosSeveridades = "tbParametrosSeveridades" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaSchema = "tbSchema" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaExcelClientes = "tbClientes" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaExcelProdutos = "tbProdutos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaExcelVendas = "tbVendas" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaExcelDadosExternos = "tbDadosExternos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared nrmExcelProdutos = let

    // Obtém os dados preparados na camada Stage.
    Fonte =
        stgExcelProdutos,

    // Descarta registros com ocorrências bloqueantes.
    RegistrosValidos =
        fxSchemaDescartarRegistrosBloqueantes(
            Fonte
        ),

    // Mantém apenas um registro para cada código de produto.
    RegistrosUnicos =
        Table.Distinct(
            RegistrosValidos,
            {"Código"}
        )

in

    RegistrosUnicos;

shared nrmExcelVendas = let

    Fonte =
        stgExcelVendas,

    RegistrosValidos =
        fxSchemaDescartarRegistrosBloqueantes(
            Fonte
        ),

    Clientes =
        List.Buffer(
            nrmExcelClientes[CPF]
        ),

    Produtos =
        List.Buffer(
            nrmExcelProdutos[Código]
        ),

    RegistrosRelacionados =
        Table.SelectRows(
            RegistrosValidos,
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

shared dimExcelProdutos = let
    // Obtém a entidade normalizada.
    Fonte =
        nrmExcelProdutos,

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

shared fxColunasTipo = (
    Tabela as table,
    Tipos as list
)
as list =>

let
    TipoTabela =
        Type.TableRow(
            Value.Type(Tabela)
        ),

    Campos =
        Type.RecordFields(
            TipoTabela
        ),

    Colunas =
        Record.FieldNames(
            Campos
        ),

    Resultado =
        List.Select(
            Colunas,
            (Coluna) =>
                List.AnyTrue(
                    List.Transform(
                        Tipos,
                        (Tipo) =>
                            Type.Is(
                                Record.Field(Campos, Coluna)[Type],
                                Tipo
                            )
                    )
                )
        )

in
    Resultado;

shared cfgIntervalosFatos = // Estrutura utilizada para descobrir o intervalo de datas
// que será utilizado na dimensão de calendario.

{
    // Adicione uma entrada por tabela fato e sua coluna de data.
    fxIntervaloData(fatoExcelVendas, "Data")
};

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

shared fxSchemaColunas = (
    tabela as text
)
as list =>

Record.FieldNames(
    fxSchema(tabela)
);

shared fxSchemaTipos = (
    tabela as text,
    optional tipo as nullable type
)
as list =>

let
    Schema =
        fxSchema(tabela),

    Colunas =
        Record.FieldNames(Schema),

    Filtradas =
        if tipo = null then
            Colunas
        else
            List.Select(
                Colunas,
                each Record.Field(Schema, _)[Tipo] = tipo
            ),

    Resultado =
        List.Transform(
            Filtradas,
            each
                {
                    _,
                    Record.Field(Schema, _)[Tipo]
                }
        )

in
    Resultado;

shared fxSchemaObrigatorias = (
    tabela as text
)
as list =>

let
    Schema =
        fxSchema(tabela),

    Resultado =
        List.Select(
            Record.FieldNames(Schema),
            (Coluna) =>
                Record.Field(
                    Schema,
                    Coluna
                )[Obrigatório]
        )

in
    Resultado;

shared fxSchemaOrdem = (
    tabela as text
)
as list =>

let
    Schema =
        fxSchema(tabela),

    Colunas =
        Table.FromRecords(
            List.Transform(
                Record.FieldNames(Schema),
                (Coluna) =>
                    [
                        Coluna = Coluna,
                        Ordem =
                            if Record.Field(Schema, Coluna)[Ordem] = null then
                                999999
                            else
                                Record.Field(Schema, Coluna)[Ordem]
                    ]
            )
        ),


    Ordenado =
        Table.Sort(
            Colunas,
            {
                {"Ordem", Order.Ascending},
                {"Coluna", Order.Ascending}
            }
        )

in
    Ordenado[Coluna];

shared fxSchemaGarantirColunas = (
    tabela as table,
    schema as text
)
as table =>

let
    ColunasSchema =
        fxSchemaColunas(schema),

    ColunasExistentes =
        Table.ColumnNames(tabela),

    ColunasFaltantes =
        List.Difference(
            ColunasSchema,
            ColunasExistentes
        ),

    // MUDANÇA DE ABORDAGEM: 
    // Em vez de utilizar a função Table.AddColumn repetidamente dentro do loop List.Accumulate,
    // o que reconstrói a estrutura da tabela em N camadas aninhadas gerando lentidão quadrática,
    // criamos uma tabela vazia temporária contendo todas as colunas faltantes (#table(ColunasFaltantes, {}))
    // e a unimos à tabela principal usando Table.Combine. Isso realiza a adição de todas as colunas faltantes
    // preenchidas com nulo em um único passo vetorizado e de alta performance.
    Resultado =
        if List.IsEmpty(ColunasFaltantes) then
            tabela
        else
            Table.Combine({tabela, #table(ColunasFaltantes, {})})

in
    Resultado
;

shared fxSchemaAplicarTipos = (
    tabela as table,
    schema as text
)
as table =>

let
    Tipos =
        fxSchemaTipos(schema),

    // MUDANÇA DE ABORDAGEM:
    // A implementação antiga aplicava indiscriminadamente 'try fxConversor(...) otherwise null' para cada
    // célula, mesmo se o valor já fosse nulo ou já possuísse o tipo desejado. Isso causava sobrecarga de exceções.
    // Agora, inserimos uma cláusula de guarda: se o valor for null, retorna null imediatamente. Se o tipo do valor
    // já corresponder ao tipo destino da coluna, retorna o próprio valor diretamente. A conversão pesada e o 
    // tratamento 'try/otherwise' só são acionados em tempo de execução para valores incompatíveis que realmente
    // precisam de transformação.
    Transformacoes =
        List.Transform(
            Tipos,
            (t) =>
                let
                    nomeColuna = t{0},
                    tipoDestino = t{1}
                in
                    {
                        nomeColuna,
                        (v) => 
                            if v = null then 
                                null 
                            else if Value.Type(v) = tipoDestino then 
                                v 
                            else 
                                try fxConversor(v, tipoDestino) otherwise null,
                        tipoDestino
                    }
        ),

    Resultado =
        Table.TransformColumns(
            tabela,
            Transformacoes,
            null,
            MissingField.Ignore
        )

in
    Resultado;

shared fxSchemaOrdenarColunas = (
    tabela as table,
    schema as text
)
as table =>

let
    ColunasSchema =
        fxSchemaOrdem(schema),

    ColunasExistentes =
        Table.ColumnNames(tabela),

    ColunasRestantes =
        List.Difference(
            ColunasExistentes,
            ColunasSchema
        ),

    NovaOrdem =
        List.Combine(
            {
                ColunasSchema,
                ColunasRestantes
            }
        ),

    Resultado =
        Table.ReorderColumns(
            tabela,
            NovaOrdem,
            MissingField.Ignore
        )

in
    Resultado;

shared fxSchemaRemoverColunas = (
    tabela as table,
    schema as text,
    optional remover as nullable logical
)
as table =>

let
    Remover =
        if remover = null then
            true
        else
            remover,

    Resultado =
        if not Remover then

            tabela

        else

            Table.SelectColumns(
                tabela,
                fxSchemaColunas(schema),
                MissingField.Ignore
            )

in
    Resultado;

shared fxSchemaAplicar = (
    tabela as table,
    schema as text,
    optional ignorarColunas as nullable logical
)
as table =>

let

    GarantirColunas =
        fxSchemaGarantirColunas(
            tabela,
            schema
        ),

    RemoverColunas =
        fxSchemaRemoverColunas(
            GarantirColunas,
            schema,
            ignorarColunas
        ),

    Processar =
        fxSchemaProcessar(
            RemoverColunas,
            schema
        ),

    AplicarTipos =
        fxSchemaAplicarTipos(
            Processar,
            schema
        ),

    OrdenarColunas =
        fxSchemaOrdenarColunas(
            AplicarTipos,
            schema
        )

in

    OrdenarColunas;

shared fxStgRemoverVazios = // Preparação física da tabela
(
    tabela as table,
    optional ignorarColunas as nullable list
) as table =>

let
    ColunasIgnoradas = if ignorarColunas = null then {} else ignorarColunas,
    ColunasAnalise = List.Difference(Table.ColumnNames(tabela), ColunasIgnoradas),
    // OTIMIZAÇÃO: Filtra linhas totalmente vazias de forma direta e vetorizada
    LinhasValidas = Table.SelectRows(
        tabela,
        each not List.IsEmpty(
            List.RemoveMatchingItems(
                Record.FieldValues(Record.SelectFields(_, ColunasAnalise, MissingField.Ignore)),
                {"", null}
            )
        )
    )
in
    LinhasValidas

;

shared fxStgPadronizar = // Higienização genérica dos dados
(
    tabela as table,
    optional schema as nullable text,
    optional colunas as nullable list
) as table =>

let
    Colunas =
        if colunas <> null then colunas
        else if schema <> null then fxSchemaTipos(schema, type text)
        else null,
    Resultado = fxTextoTrim(tabela, Colunas)
in
    Resultado
;

shared fxStgAplicar = (
    tabela as table,
    schema as text,
    optional ignorarColunas as nullable logical
)
as table =>

let
    Preparada =
        fxStgRemoverVazios(
            tabela,
            ignorarColunas
        ),

    Estruturada =
        fxSchemaAplicar(
            Preparada,
            schema,
            ignorarColunas
        ),

    Padronizada =
        fxStgPadronizar(
            Estruturada,
            schema
        )

in
    Padronizada;

shared fxTextoTrim = // ------------------------------------------------------------------------------
// 1. fxTextoTrim (Otimizada: Sem Exceções via 'is text')
// ------------------------------------------------------------------------------
(
    tabela as table,
    optional colunas as nullable list
) as table =>

let
    TodasColunas = Table.ColumnNames(tabela),
    ColunasParaTratar = if colunas = null then TodasColunas else List.Intersect({colunas, TodasColunas}),
    // OTIMIZAÇÃO CRÍTICA: Substituição do try/otherwise por checagem estática 'is text'
    Resultado = Table.TransformColumns(
        tabela,
        List.Transform(
            ColunasParaTratar,
            (col) => {col, (v) => if v is text then Text.Trim(v) else v, type any}
        )
    )
in
    Resultado;

shared fxSchemaAplicarTratamentos = (
    valores as list,
    tratamentos as list
)
as list =>

let

    Resultado =

        List.Accumulate(

            tratamentos,

            valores,

            (Estado, Etapa) =>

                let

                    Funcao =
                        Etapa[Definição][Função],

                    Parametros =
                        Record.FieldOrDefault(
                            Etapa,
                            "Parametros"
                        )

                in

                    List.Transform(

                        Estado,

                        each

                            Funcao(

                                _,

                                Parametros

                            )

                    )

        )

in

    Resultado;

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

    Definicao =
        Record.FieldOrDefault(
            contexto,
            "Definicao",
            []
        ),

    Etapa =
        Record.FieldOrDefault(
            contexto,
            "Etapa",
            []
        ),

    Resultado =

    [

        Codigo =
            codigo,

        Severidade =
            Record.FieldOrDefault(
                contexto,
                "Severidade"
            ),

        Etapa =
            Record.FieldOrDefault(
                Etapa[Definição],
                "Código"
            ),

        Tabela =
            Record.FieldOrDefault(
                contexto,
                "Tabela"
            ),

        Coluna =
            Record.FieldOrDefault(
                contexto,
                "Coluna"
            ),

        ValorOriginal =
            valorOriginal,

        Valor =
            valor,

        Tipo =
            Record.FieldOrDefault(
                Definicao,
                "Tipo"
            ),

        Obrigatorio =
            Record.FieldOrDefault(
                Definicao,
                "Obrigatório"
            ),

        Mensagem =
            descricao,

        Detalhes =
            detalhes

    ]

in

    Resultado;

shared fxSchemaProcessar = (
    tabela as table,
    schema as text
) as table =>

let
    Definicoes = fxSchema(schema),

    NomesTabela = Table.ColumnNames(tabela),

    TemItens = (regras as any) as logical =>
        regras <> null
        and Value.Is(regras, type list)
        and not List.IsEmpty(regras),

    ColunasProcessaveis =
        List.Select(
            Record.FieldNames(Definicoes),
            (coluna as text) =>
                let
                    Definicao = Record.Field(Definicoes, coluna),
                    Tratamentos =
                        Record.FieldOrDefault(
                            Definicao,
                            "Tratamentos"
                        ),
                    Validacoes =
                        Record.FieldOrDefault(
                            Definicao,
                            "Validações"
                        )
                in
                    List.Contains(NomesTabela, coluna)
                    and (
                        TemItens(Tratamentos)
                        or TemItens(Validacoes)
                    )
        ),

    Resultado =
        if List.IsEmpty(ColunasProcessaveis) then
            tabela
        else
            let
                EstadoInicial =
                    [
                        NomesColunas = NomesTabela,
                        Colunas = Table.ToColumns(tabela)
                    ],

                EstadoFinal =
                    List.Accumulate(
                        ColunasProcessaveis,
                        EstadoInicial,
                        (estado as record, coluna as text) =>
                            fxSchemaProcessarColuna(
                                estado,
                                schema,
                                coluna,
                                Record.Field(
                                    Definicoes,
                                    coluna
                                )
                            )
                    )
            in
                Table.FromColumns(
                    EstadoFinal[Colunas],
                    EstadoFinal[NomesColunas]
                )
in
    Resultado
;

shared fxSchemaAplicarValidacoes = (
    valores as list,
    validacoes as list,
    contexto as record
)
as list =>

let

    EstadoInicial =

        List.Transform(

            valores,

            each

                [

                    Valor =
                        _,

                    Ocorrencias =
                        null

                ]

        ),

    Resultado =

        List.Accumulate(

            validacoes,

            EstadoInicial,

            (Estado, Etapa) =>

                let

                    ContextoEtapa =

                        contexto &

                        [

                            Etapa =
                                Etapa,

                            Severidade =
                                Record.Field(

                                    cfgParametrosSeveridades,

                                    Etapa[Definição][Severidade]

                                )

                        ]

                in

                    List.Transform(

                        Estado,

                        (Item) =>

                            if Item[Ocorrencias] <> null then

                                Item

                            else

                                let

                                    ResultadoValidacao =

                                        Function.Invoke(

                                            Etapa[Definição][Função],

                                            {

                                                Item[Valor],

                                                Record.FieldOrDefault(
                                                    Etapa,
                                                    "Parametros"
                                                ),

                                                ContextoEtapa

                                            }

                                        ),

                                    NovasOcorrencias =
                                        ResultadoValidacao[Ocorrencias] ?? {},

                                    Ocorrencias =

                                        if List.IsEmpty(NovasOcorrencias) then

                                            null

                                        else

                                            NovasOcorrencias

                                in

                                    [

                                        Valor =
                                            ResultadoValidacao[Valor],

                                        Ocorrencias =
                                            Ocorrencias

                                    ]

                    )

        )

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
                "'. Utilize o formato NOME{parametro1,parametro2}.",

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

                        Permitidos = parametros

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
        Text.Upper(
            Text.From(valor)
        ),

    Valido =
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
        Text.Length(
            Text.From(valor)
        ),

    Valido =
        TamanhoEsperado <> null
        and TamanhoAtual = TamanhoEsperado,

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
        if parametros = null or List.IsEmpty(parametros) then
            null
        else
            Number.From(parametros{0}),

    Valor =
        Number.From(valor),

    Valido =
        Minimo <> null
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
        if parametros = null or List.IsEmpty(parametros) then
            null
        else
            Number.From(parametros{0}),

    Valor =
        Number.From(valor),

    Valido =
        Maximo <> null
        and Valor <= Maximo,

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
        if parametros = null or List.Count(parametros) < 1 then
            null
        else
            Number.From(parametros{0}),

    Maximo =
        if parametros = null or List.Count(parametros) < 2 then
            null
        else
            Number.From(parametros{1}),

    Valor =
        Number.From(valor),

    Valido =
        Minimo <> null
        and Maximo <> null
        and Valor >= Minimo
        and Valor <= Maximo,

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

    Texto =
        Text.From(valor),

    Partes =
        Text.Split(
            Texto,
            "@"
        ),

    Valido =
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
        Text.Trim(
            Text.From(valor)
        ),

    Protocolo =
        if Text.StartsWith(URL, "http://") then
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
        Protocolo <> null
        and Text.Contains(Dominio, ".")
        and Text.PositionOf(Dominio, ".") > 0
        and Text.PositionOf(Dominio, ".") < Text.Length(Dominio) - 1,

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
        FormatoComMascara,

    Ocorrencias =
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
        if not NaoRepetido then

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

    CNPJ =
        Text.Select(
            Text.From(valor),
            {"0".."9"}
        ),

    Digitos =
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
                    each Digitos{_} * PesosDV1{_}
                )
            )
        else
            null,

    RestoDV1 =
        if NaoRepetido then
            Number.Mod(SomaDV1, 11)
        else
            null,

    DV1 =
        if RestoDV1 < 2 then
            0
        else
            11 - RestoDV1,

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
                        * PesosDV2{_}
                )
            )
        else
            null,

    RestoDV2 =
        if NaoRepetido then
            Number.Mod(SomaDV2, 11)
        else
            null,

    DV2 =
        if RestoDV2 < 2 then
            0
        else
            11 - RestoDV2,

    Valido =
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

shared tstSchema = let
    Fonte =
        fxSchema("tbClientes")
in
    Fonte;

shared fxSchemaProcessarColuna = (
    estado as record,
    schema as text,
    coluna as text,
    definicao as record
)
as record =>

let

    ProcessarTratamentos =
        fxParametro(
            "Processar_Tratamentos"
        ),

    ProcessarValidacoes =
        fxParametro(
            "Processar_Validacoes"
        ),

    NomesColunas =
        estado[NomesColunas],

    Colunas =
        estado[Colunas],

    Posicao =
        List.PositionOf(
            NomesColunas,
            coluna
        ),

    Tratamentos =
        Record.FieldOrDefault(
            definicao,
            "Tratamentos"
        ),

    Validacoes =
        if ProcessarValidacoes then
            Record.FieldOrDefault(
                definicao,
                "Validações"
            )
        else
            {},

    TemTratamentos =
        ProcessarTratamentos
        and
        Tratamentos <> null
        and
        not List.IsEmpty(
            Tratamentos
        ),

    TemValidacoes =
        Validacoes <> null
        and
        not List.IsEmpty(
            Validacoes
        ),

    ValoresOriginais =
        Colunas{Posicao},

    ValoresTratados =

        if TemTratamentos then

            fxSchemaAplicarTratamentos(
                ValoresOriginais,
                Tratamentos
            )

        else

            ValoresOriginais,

    ResultadoValidacao =

        if TemValidacoes then

            fxSchemaAplicarValidacoes(

                ValoresTratados,

                Validacoes,

                [

                    Definicao =
                        definicao,

                    Tabela =
                        schema,

                    Coluna =
                        coluna

                ]

            )

        else

            null,

    ValoresFinais =

        if TemValidacoes then

            List.Transform(
                ResultadoValidacao,
                each [Valor]
            )

        else

            ValoresTratados,

    ColunasAtualizadas =

        if
            TemTratamentos
            or
            TemValidacoes
        then

            List.ReplaceRange(

                Colunas,

                Posicao,

                1,

                {
                    ValoresFinais
                }

            )

        else

            Colunas,

    PosicaoOcorrencias =

        if TemValidacoes then

            List.PositionOf(
                NomesColunas,
                "Ocorrencias"
            )

        else

            -1,

    OcorrenciasNovas =

        if TemValidacoes then

            List.Transform(
                ResultadoValidacao,
                each [Ocorrencias] ?? {}
            )

        else

            null,

    OcorrenciasConsolidadas =

        if
            not TemValidacoes
        then

            null

        else

            List.Transform(

                List.Positions(
                    OcorrenciasNovas
                ),

                (i) =>

                    let

                        Existentes =

                            if PosicaoOcorrencias >= 0 then

                                ColunasAtualizadas{PosicaoOcorrencias}{i} ?? {}

                            else

                                {},

                        Novas =
                            OcorrenciasNovas{i},

                        Todas =
                            List.Combine(
                                {
                                    Existentes,
                                    Novas
                                }
                            )

                    in

                        if List.IsEmpty(Todas) then
                            null
                        else
                            Todas

            ),

    NovasColunas =

        if
            not TemValidacoes
        then

            ColunasAtualizadas

        else if
            PosicaoOcorrencias >= 0
        then

            List.ReplaceRange(

                ColunasAtualizadas,

                PosicaoOcorrencias,

                1,

                {
                    OcorrenciasConsolidadas
                }

            )

        else

            ColunasAtualizadas
            &
            {
                OcorrenciasConsolidadas
            },

    NovosNomes =

        if
            not TemValidacoes
            or
            PosicaoOcorrencias >= 0
        then

            NomesColunas

        else

            NomesColunas
            &
            {
                "Ocorrencias"
            }

in

    [

        NomesColunas =
            NovosNomes,

        Colunas =
            NovasColunas

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
];

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

        valor = null

        or

        (
            valor is text
            and
            Text.Trim(valor) = ""
        ),

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

shared benchCustoMotorSTG = let
    Repeticoes =
        10,

    Consultas =
        Table.SelectRows(
            diagConsultasPQ,
            each
                [Categoria] = "Stage"
                and Value.Is([Valor], type table)
        ),

    Benchmark =
        Table.AddColumn(
            Consultas,
            "Benchmark",
            (Consulta) =>

                let
                    Tabela =
                        Consulta[Valor],

                    Inicio =
                        DateTimeZone.UtcNow(),

                    Contagens =
                        List.Transform(
                            {1..Repeticoes},
                            each
                                Table.RowCount(
                                    Table.Buffer(Tabela)
                                )
                        ),

                    Fim =
                        DateTimeZone.UtcNow(),

                    Tempo =
                        Duration.TotalSeconds(
                            Fim - Inicio
                        )

                in
                    [
                        Linhas = List.Last(Contagens),
                        Execuções = Repeticoes,
                        #"Tempo Total (s)" = Tempo,
                        #"Tempo Médio (s)" = Tempo / Repeticoes
                    ],
            type record
        ),

    Expandido =
        Table.ExpandRecordColumn(
            Benchmark,
            "Benchmark",
            {
                "Linhas",
                "Execuções",
                "Tempo Total (s)",
                "Tempo Médio (s)"
            }
        ),

    Resultado =
        Table.SelectColumns(
            Expandido,
            {
                "Consulta",
                "Linhas",
                "Execuções",
                "Tempo Total (s)",
                "Tempo Médio (s)"
            }
        )

in
    Resultado;

shared fxSchemaDescartarRegistrosBloqueantes = (
    tabela as table
)
as table =>

let

    ProcessarTratamentos =
        fxParametro("Processar_Tratamentos"),

    ProcessarValidacoes =
        fxParametro("Processar_Validacoes"),

    DescartarRegistrosBloqueantes =
        fxParametro("Descartar_Registros_Bloqueantes"),

    TemOcorrencias =
        Table.HasColumns(
            tabela,
            "Ocorrencias"
        ),

    TabelaFiltrada =

        if
            not DescartarRegistrosBloqueantes
            or
            not ProcessarValidacoes
            or
            not TemOcorrencias
        then

            tabela

        else

            Table.SelectRows(
                tabela,

                each

                    let

                        Ocorrencias =

                            if
                                Value.Is(
                                    [Ocorrencias],
                                    type list
                                )
                            then
                                [Ocorrencias]
                            else
                                {},

                        TemBloqueio =

                            List.AnyTrue(

                                List.Transform(

                                    Ocorrencias,

                                    each

                                        Record.FieldOrDefault(

                                            _[Severidade],

                                            "Bloqueia",

                                            false

                                        )

                                )

                            )

                    in

                        not TemBloqueio

            ),

    Resultado =

        if
            TemOcorrencias
        then
            Table.RemoveColumns(
                TabelaFiltrada,
                "Ocorrencias"
            )
        else
            TabelaFiltrada

in

    Resultado;