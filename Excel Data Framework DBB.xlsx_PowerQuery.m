// Power Query from: Excel Data Framework DBB.xlsx
// Pathname: c:\Users\daniel-bighelini\OneDrive\Documentos\Planilhas\Excel Data Framework DBB\Excel Data Framework DBB.xlsx
// Extracted: 2026-07-24T03:43:27.660Z

section Section1;

shared srcCategoriasPowerQuery = let
    Fonte = srcWorkbook{[Name=parTabelaCategoriasConsultasPQ]}[Content]
in
    Fonte;

shared cfgSections = let

    Fonte = #sections[Section1],
    
    Objetos =

        Record.SelectFields(

            Fonte,

            List.Select(

                Record.FieldNames(Fonte),

                each
                    not Text.StartsWith(_, "cfg") and
                    not Text.StartsWith(_, "diag")

            )

        )

in
    Objetos;

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

shared parTabelaParametrosBooleanos = "tbParametrosBooleanos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaParametrosTipos = "tbParametrosTipos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared cfgParametros = // Materializa a tabela consolidada em um Record para permitir
// acesso O(1) aos parâmetros através de Record.Field.
Record.FromTable(
    Table.Buffer(stgParametros)
);

shared cfgParametrosCategoriasPowerQuery = let
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

shared cfgParametrosPowerQuery = let

    Fonte =
        cfgSections,

    NomesCandidatos =

        List.Select(

            Record.FieldNames(Fonte),

            each
                Text.StartsWith(_, "par")

        ),

    ConsultasCandidatas =

        Record.SelectFields(

            Fonte,

            NomesCandidatos,

            MissingField.Ignore

        ),

    NomesParametros =

        List.Select(

            Record.FieldNames(ConsultasCandidatas),

            each

                fxParametrosIdentificarParametroPQ(

                    Record.Field(

                        ConsultasCandidatas,

                        _

                    )

                )

        ),

    Parametros =

        Record.SelectFields(

            ConsultasCandidatas,

            NomesParametros,

            MissingField.Ignore

        )

in

    Parametros;

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

    RemoverVazios =
        fxSchemaRemoverVaziosOLD(
            Fonte
        ),

    Padronizado =
        fxSchemaPadronizarOLD(
            RemoverVazios
        ),

    Distintos =
        Table.Distinct(
            Padronizado,
            {"Prefixo"}
        ),
    TipoAlterado = Table.TransformColumnTypes(Distintos,{{"Ordem", Int64.Type}, {"Prefixo", type text}, {"Categoria", type text}, {"Objetivo", type text}, {"Saída", type text}})

in

    TipoAlterado;

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

    Fonte =
        srcParametrosTratamentos,

    RemoverVazios =
        fxSchemaRemoverVaziosOLD(
            Fonte,
            {"Ativo"}
        ),

    Padronizado =
        fxSchemaPadronizarOLD(
            RemoverVazios
        ),

    TratamentosAtivos =
        Table.SelectRows(
            Padronizado,
            each [Ativo] = true
        ),

    Funcoes =
        cfgSections,

    Tratamentos =
        Table.TransformColumns(
            TratamentosAtivos,
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
                    "TipoEntrada",
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
                    "TipoSaída",
                    each
                        if _ = null then
                            null
                        else
                            Text.Trim(
                                Text.From(_)
                            ),
                    type nullable text
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

    TratamentosUnicos

;

shared stgParametrosValidacoes = let

    Fonte =
        srcParametrosValidacoes,

    RemoverVazios =
        fxSchemaRemoverVaziosOLD(
            Fonte,
            {"Ativo"}
        ),

    Padronizado =
        fxSchemaPadronizarOLD(
            RemoverVazios
        ),

    ValidacoesAtivos =
        Table.SelectRows(
            Padronizado,
            each [Ativo] = true
        ),

    Funcoes =
        cfgSections,

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
                },
                {
                    "TipoEntrada",
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
                    "TipoSaída",
                    each
                        if _ = null then
                            null
                        else
                            Text.Trim(
                                Text.From(_)
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

    RemoverVazios =
        fxSchemaRemoverVaziosOLD(
            Fonte
        ),

    Padronizado =
        fxSchemaPadronizarOLD(
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

                        Consulta = [Name],

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

    RegistrosValidos =
        fxSchemaDescartarRegistrosBloqueantes(
            Fonte
        )

in

    RegistrosValidos;

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

    Fonte =
        Table.Buffer(
            stgParametrosTratamentos
        ),

    Nomes =
        Table.Column(
            Fonte,
            "Name"
        ),

    Valores =
        Table.AddColumn(
            Fonte,
            "RecordValue",
            each
                [
                    Código = [Código],
                    Função = [Função],
                    TipoEntrada = [TipoEntrada],
                    TipoSaída = [TipoSaída],
                    Padrão = [Padrão],
                    Ativo = [Ativo]
                ],
            type record
        )[RecordValue],

    Resultado =
        Record.FromList(
            Valores,
            Nomes
        )

in

    Resultado;

shared cfgParametrosTratamentosGenericos = let

    Codigos =
    {
        "TRIM",
        "CLEAN",
        "SINGLESPACE",
        "EMPTYTONULL"
    },

    Tratamentos =
        Record.SelectFields(
            cfgParametrosTratamentos,
            Codigos,
            MissingField.Ignore
        )

in

    Tratamentos;

shared cfgParametrosValidacoes = let

    Fonte =
        Table.Buffer(
            stgParametrosValidacoes
        ),

    Nomes =
        Table.Column(
            Fonte,
            "Name"
        ),

    Valores =
        Table.AddColumn(
            Fonte,
            "RecordValue",
            each
                [
                    Código = [Código],
                    Função = [Função],
                    Severidade = [Severidade],
                    TipoEntrada = [TipoEntrada],
                    TipoSaída = [TipoSaída],
                    Padrão = [Padrão],
                    Ativo = [Ativo]
                ],
            type record
        )[RecordValue],

    Resultado =
        Record.FromList(
            Valores,
            Nomes
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

shared parTabelaClientes = "tbClientes" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaProdutos = "tbProdutos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaVendas = "tbVendas" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared parTabelaDadosGenericos = "tbDadosGenericos" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared nrmProdutos = let

    // Obtém os dados preparados na camada Stage.
    Fonte =
        stgProdutos,

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

shared nrmVendas = let

    Fonte =
        stgVendas,

    RegistrosValidos =
        fxSchemaDescartarRegistrosBloqueantes(
            Fonte
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

shared cfgIntervalosFatos = // Estrutura utilizada para descobrir o intervalo de datas
// que será utilizado na dimensão de calendario.

{
    // Adicione uma entrada por tabela fato e sua coluna de data.
    fxIntervaloData(fatoVendas, "Data")
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
    schema as text
)
as list =>

Record.FieldNames(
    fxSchema(schema)
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

shared fxSchemaRemoverVaziosOLD = // Preparação física da tabela
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

shared fxSchemaPadronizarOLD = /* Higienização genérica dos dados
Responsabilidades:
- Operações universais, independentes do domínio.
*/

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
    Resultado;

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

shared fxSchemaProcessarTratamentos = (
    estado as record,
    schema as text,
    coluna as text,
    definicao as record
)
as record =>

let

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

    TemTratamentos =
        Tratamentos <> null
        and
        not List.IsEmpty(
            Tratamentos
        ),

    ValoresOriginais =
        Colunas{Posicao},

    ValoresFinais =

        if TemTratamentos then

            fxSchemaAplicarTratamentos(
                ValoresOriginais,
                Tratamentos
            )

        else

            ValoresOriginais,

    NovasColunas =

        if TemTratamentos then

            List.ReplaceRange(

                Colunas,

                Posicao,

                1,

                {
                    ValoresFinais
                }

            )

        else

            Colunas

in

    [

        NomesColunas =
            NomesColunas,

        Colunas =
            NovasColunas

    ];

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
)
as table =>

let

    Definicoes =
        fxSchema(schema),

    NomesTabela =
        Table.ColumnNames(tabela),

    ProcessarTratamentos =
        fxParametro(
            "Processar_Tratamentos"
        ),

    ProcessarValidacoes =
        fxParametro(
            "Processar_Validacoes"
        ),

    TemItens =

        (regras as any) as logical =>

            regras <> null
            and
            Value.Is(
                regras,
                type list
            )
            and
            not List.IsEmpty(
                regras
            ),

    ColunasTratamentos =

        if ProcessarTratamentos then

            List.Select(

                Record.FieldNames(
                    Definicoes
                ),

                (coluna as text) =>

                    let

                        Definicao =
                            Record.Field(
                                Definicoes,
                                coluna
                            )

                    in

                        List.Contains(
                            NomesTabela,
                            coluna
                        )
                        and
                        TemItens(
                            Record.FieldOrDefault(
                                Definicao,
                                "Tratamentos"
                            )
                        )

            )

        else

            {},

    ColunasValidacoes =

        if ProcessarValidacoes then

            List.Select(

                Record.FieldNames(
                    Definicoes
                ),

                (coluna as text) =>

                    let

                        Definicao =
                            Record.Field(
                                Definicoes,
                                coluna
                            )

                    in

                        List.Contains(
                            NomesTabela,
                            coluna
                        )
                        and
                        TemItens(
                            Record.FieldOrDefault(
                                Definicao,
                                "Validações"
                            )
                        )

            )

        else

            {},

    EstadoInicial =

        [

            NomesColunas =
                NomesTabela,

            Colunas =
                Table.ToColumns(
                    tabela
                )

        ],

    EstadoTratamentos =

        List.Accumulate(

            ColunasTratamentos,

            EstadoInicial,

            (estado as record, coluna as text) =>

                fxSchemaProcessarTratamentos(

                    estado,

                    schema,

                    coluna,

                    Record.Field(
                        Definicoes,
                        coluna
                    )

                )

        ),

    EstadoValidacoes =

        List.Accumulate(

            ColunasValidacoes,

            EstadoTratamentos,

            (estado as record, coluna as text) =>

                fxSchemaProcessarValidacoes(

                    estado,

                    schema,

                    coluna,

                    Record.Field(
                        Definicoes,
                        coluna
                    )

                )

        ),

    Resultado =

        Table.FromColumns(

            EstadoValidacoes[Colunas],

            EstadoValidacoes[NomesColunas]

        )

in

    Resultado;

shared fxSchemaProcessarValidacoes = (
    estado as record,
    schema as text,
    coluna as text,
    definicao as record
)
as record =>

let

    NomesColunas =
        estado[NomesColunas],

    Colunas =
        estado[Colunas],

    Posicao =
        List.PositionOf(
            NomesColunas,
            coluna
        ),

    Validacoes =
        Record.FieldOrDefault(
            definicao,
            "Validações"
        ),

    TemValidacoes =
        Validacoes <> null
        and
        not List.IsEmpty(
            Validacoes
        ),

    ValoresOriginais =
        Colunas{Posicao},

    ResultadoValidacao =

        if TemValidacoes then

            fxSchemaAplicarValidacoes(

                ValoresOriginais,

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

            ValoresOriginais,

    ColunasAtualizadas =

        if TemValidacoes then

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

        if not TemValidacoes then

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

                    Funcao =
                        Etapa[Definição][Função],

                    Parametros =
                        Record.FieldOrDefault(
                            Etapa,
                            "Parametros"
                        ),

                    Severidade =
                        Record.Field(
                            cfgParametrosSeveridades,
                            Etapa[Definição][Severidade]
                        ),

                    ContextoEtapa =

                        contexto &

                        [

                            Etapa =
                                Etapa,

                            Severidade = Severidade

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

                                        Funcao(

                                            Item[Valor],

                                            Parametros,

                                            ContextoEtapa

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

shared tstGeral = let

    Schema = cfgSchema[tbClientes]
in
    Schema;

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

shared parTabelaCategoriasConsultasPQ = "tbSobreCategoriasConsultasPQ" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true];

shared fxStgPreparar = /*
Responsabilidades:
- Remover linhas completamente vazias.
- Garantir que a tabela possua pelo menos uma coluna.
- Garantir que os nomes das colunas sejam válidos:
    - não nulos;
    - não vazios;
    - únicos.
- Normalizar os nomes das colunas removendo espaços nas extremidades.
- Nunca alterar os valores das células.
- Nunca depender do Schema.
*/

(
    tabela as table
)
as table =>

let

    // Remove registros completamente vazios.
    TabelaSemLinhasVazias =
        Table.SelectRows(
            tabela,
            each
                List.AnyTrue(
                    List.Transform(
                        Record.FieldValues(_),
                        each _ <> null and _ <> ""
                    )
                )
        ),

    // Obtém os nomes atuais das colunas.
    Colunas =
        Table.ColumnNames(TabelaSemLinhasVazias),

    // Remove espaços nas extremidades dos nomes.
    ColunasNormalizadas =
        List.Transform(
            Colunas,
            each Text.Trim(_)
        ),

    // Verifica nomes vazios.
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
        List.Count(List.Distinct(ColunasNormalizadas)),

    _ErroDuplicidade =
        if TemDuplicidade then
            error "A tabela possui nomes de colunas duplicados."
        else
            null,

    // Garante que exista pelo menos uma coluna.
    _ErroSemColunas =
        if List.IsEmpty(ColunasNormalizadas) then
            error "A tabela não possui colunas."
        else
            null,

    Resultado =
        Table.RenameColumns(
            TabelaSemLinhasVazias,
            List.Zip(
                {
                    Colunas,
                    ColunasNormalizadas
                }
            ),
            MissingField.Ignore
        )

in

    Resultado;

shared fxStgPadronizar = /*
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

    EstadoInicial =

        [

            NomesColunas =
                Table.ColumnNames(
                    tabela
                ),

            Colunas =
                Table.ToColumns(
                    tabela
                )

        ],

    EstadoFinal =

        List.Accumulate(

            Record.FieldNames(
                pipeline
            ),

            EstadoInicial,

            (
                estado,
                coluna
            ) =>

                let

                    PipelineColuna =

                        Record.Field(
                            pipeline,
                            coluna
                        ),

                    Tratamentos =

                        PipelineColuna[Tratamentos],

                    Posicao =

                        List.PositionOf(
                            estado[NomesColunas],
                            coluna
                        ),

                    Valores =

                        estado[Colunas]{Posicao},

                    NovosValores =

                        if Tratamentos = null then

                            Valores

                        else

                            fxStgAplicarTratamentos(
                                Valores,
                                Tratamentos
                            ),

                    NovasColunas =

                        if Tratamentos = null then

                            estado[Colunas]

                        else

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

                        NomesColunas =
                            estado[NomesColunas],

                        Colunas =
                            NovasColunas

                    ]

        ),

    Resultado =

        Table.FromColumns(

            EstadoFinal[Colunas],

            EstadoFinal[NomesColunas]

        )

in

    Resultado;

shared cfgTabelasExcel = /*
------------------------------------------------------------------------------
Consulta.....: cfgTabelasExcel

Descrição....:
Catálogo das tabelas disponíveis em Excel.CurrentWorkbook().

Materializa informações frequentemente utilizadas pelo framework,
permitindo acesso rápido às tabelas e seus metadados.

Estrutura:

    Source
        Origem do objeto.

    Value
        Referência para a tabela.

    Columns
        Lista das colunas da tabela.

------------------------------------------------------------------------------
*/

let

//--------------------------------------------------------------------------
// Objetos
//--------------------------------------------------------------------------

    Objetos =

        List.Select(

            Record.FieldValues(
                cfgObjetos[Excel]
            ),

            each

                [Kind] = "Table"

        ),

//--------------------------------------------------------------------------
// Tabelas
//--------------------------------------------------------------------------

    Tabelas =

        Record.FromList(

            List.Transform(

                Objetos,

                each

                    let

                        Tabela =

                            Table.SelectRows(

                                srcWorkbook,

                                (r) =>

                                    r[Name] = [Name]

                            ){0}[Content]

                    in

                        [

                            Name = [Name],

                            Source = [Source],

                            Value = Tabela,

                            Columns =

                                List.Buffer(

                                    Table.ColumnNames(
                                        Tabela
                                    )

                                )

                        ]

            ),

            List.Transform(

                Objetos,

                each [Name]

            )

        )
in
    Tabelas;

shared cfgTabelasPowerQuery = /*
------------------------------------------------------------------------------
Consulta.....: cfgTabelasPQ

Descrição....:
Catálogo das tabelas do ambiente Power Query.

Materializa informações frequentemente utilizadas pelo framework,
permitindo acesso rápido às tabelas e seus metadados.

Estrutura:

    Source
        Origem do objeto.

    Value
        Referência para a tabela.

    Columns
        Lista das colunas da tabela.

------------------------------------------------------------------------------
*/

let

//--------------------------------------------------------------------------
// Objetos
//--------------------------------------------------------------------------

    Fonte =

        Record.FieldValues(
            cfgObjetos[PowerQuery]
        ),

    Objetos =

        List.Select(

            Fonte,

            each

                [Kind] = "Table"

        ),

//--------------------------------------------------------------------------
// Tabelas
//--------------------------------------------------------------------------

    Tabelas =

        Record.FromList(

            List.Transform(

                Objetos,

                each

                    let

                        Tabela =

                            Record.Field(

                                cfgSections,

                                [Name]

                            )

                    in

                        [

                            Name = [Name],

                            Source = [Source],

                            Value = Tabela,

                            Columns =

                                List.Buffer(

                                    Table.ColumnNames(
                                        Tabela
                                    )

                                )

                        ]

            ),

            List.Transform(

                Objetos,

                each [Name]

            )

        )

in

    Tabelas;

shared cfgTabelas = let

    Modelo =

        [

            PowerQuery =

                cfgTabelasPowerQuery,

            Excel =

                cfgTabelasExcel

        ]
in
    Modelo;

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

shared cfgObjetos = /*
------------------------------------------------------------------------------
Consulta.....: cfgObjetos

Descrição....:
Catálogo central dos objetos disponíveis no ambiente.

Não materializa consultas nem funções.

Cada item contém apenas metadados suficientes para localizar e classificar
objetos do ambiente.

Estrutura:

    PowerQuery
        Record contendo os objetos do Power Query.

    Excel
        Record contendo os objetos do Excel.

------------------------------------------------------------------------------
*/

let

//--------------------------------------------------------------------------
// Objetos Power Query
//--------------------------------------------------------------------------

    ObjetosPowerQuery =

        List.RemoveNulls(

            List.Transform(

                Record.FieldNames(
                    cfgSections
                ),

                (Nome) =>

                    let

                        Metadados =

                            try
                                fxIdentificarObjetoPeloNome(
                                    Nome,
                                    "PowerQuery"
                                )

                    in

                        if Metadados[HasError] then

                            null

                        else

                            Record.Combine(

                                {

                                    [

                                        Name = Nome,

                                        Source = "PowerQuery"

                                    ],

                                    Metadados[Value]

                                }

                            )

            )

        ),

    PowerQuery =

        Record.FromList(

            ObjetosPowerQuery,

            List.Transform(

                ObjetosPowerQuery,

                each [Name]

            )

        ),

//--------------------------------------------------------------------------
// Objetos Excel
//--------------------------------------------------------------------------

    ObjetosExcel =

        List.Transform(

            Table.Column(

                srcWorkbook,

                "Name"

            ),

            (Nome) =>

                let

                    Metadados =

                        fxIdentificarObjetoPeloNome(

                            Nome,

                            "Excel"

                        )

                in

                    Record.Combine(

                        {

                            [

                                Name = Nome,

                                Source = "Excel"

                            ],

                            Metadados

                        }

                    )

        ),

    Excel =

        Record.FromList(

            ObjetosExcel,

            List.Transform(

                ObjetosExcel,

                each [Name]

            )

        ),

//--------------------------------------------------------------------------
// Resultado
//--------------------------------------------------------------------------

    Resultado =

        [

            PowerQuery = PowerQuery,

            Excel = Excel

        ]

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

let

    Fonte =

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

in

    Fonte;

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

shared cfgFuncoesPowerQuery = let

    Objetos =

        List.Select(

            Record.FieldValues(
                cfgObjetos[PowerQuery]
            ),

            each

                [Source] = "PowerQuery"
                and [Kind] = "Function"

        ),

    Funcoes =

        Record.FromList(

            List.Transform(

                Objetos,

                each
                    Record.Field(
                        cfgSections,
                        [Name]
                    )

            ),

            List.Transform(

                Objetos,

                each [Name]

            )

        )

in

    Funcoes;

shared fxStgPipelineGerar = /*
------------------------------------------------------------------------------
Função....: fxPipelineGerar

Descrição.:
Gera o pipeline completo de uma tabela do framework.

------------------------------------------------------------------------------
*/

(
    tabela as text
)
as record =>

let

    //----------------------------------------------------------------------
    // Obtém a definição da tabela
    //----------------------------------------------------------------------

    DefinicaoTabela =

        if Record.HasFields(
            cfgObjetos[PowerQuery],
            tabela
        ) then

            Record.Field(
                cfgTabelas[PowerQuery],
                tabela
            )

        else if Record.HasFields(
            cfgObjetos[Excel],
            tabela
        ) then

            Record.Field(
                cfgTabelas[Excel],
                tabela
            )

        else

            error Error.Record(
                "Tabela inválida",
                "A tabela não existe no framework.",
                [
                    Tabela = tabela
                ]
            ),

    //----------------------------------------------------------------------
    // Colunas
    //----------------------------------------------------------------------

    Colunas =

        DefinicaoTabela[Columns],

    //----------------------------------------------------------------------
    // Schema
    //----------------------------------------------------------------------

    Schema =

        Record.FieldOrDefault(
            cfgSchema,
            tabela,
            []
        ),

    //----------------------------------------------------------------------
    // Pipeline
    //----------------------------------------------------------------------

    Pipeline =

        Record.FromList(

            List.Transform(

                Colunas,

                (Coluna) =>

                    let

                        DefinicaoColuna =

                            Record.FieldOrDefault(
                                Schema,
                                Coluna,
                                []
                            )

                    in

                        Record.Combine({

                            [

                                Tipo = type any,

                                Obrigatório = false,

                                Ordem = null,

                                Tratamentos = {},

                                Validações = {}

                            ],

                            DefinicaoColuna

                        })

            ),

            Colunas

        )

in

    Pipeline;

shared fxStgAplicarTratamentos = (
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

shared fxStgAplicarTipos = /*
Responsabilidades:
- Aplicar os tipos definidos no Pipeline.
- Inferir automaticamente o tipo quando o Pipeline informar type any.
- Nunca aplicar tratamentos.
- Nunca aplicar validações.
- Nunca adicionar, remover ou ordenar colunas.
*/

(
    tabela as table,
    pipeline as record,
    optional amostra as nullable number
)
as table =>

let

    Transformacoes =

        List.Transform(

            List.Intersect(
                {
                    Table.ColumnNames(tabela),
                    Record.FieldNames(pipeline)
                }
            ),

            (Coluna) =>

                let

                    PipelineColuna =

                        Record.Field(
                            pipeline,
                            Coluna
                        ),

                    TipoConfigurado =

                        PipelineColuna[Tipo],

                    TipoFinal =

                        if TipoConfigurado = type any then

                            fxStgIdentificarTipoColuna(
                                tabela,
                                Coluna,
                                amostra
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

        List.Sort(

            Record.FieldNames(
                cfgParametrosCategoriasPowerQuery
            ),

            (x, y) =>

                Value.Compare(

                    Text.Length(y),

                    Text.Length(x)

                )

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

                cfgParametrosCategoriasPowerQuery,

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

shared tstStgPipelineGerar = let
    Fonte = fxStgPipelineGerar("tbClientes")
in
    Fonte;

shared tstSchema = fxSchema("tbClientes");

shared tstStgPadronizarNew = let

    Fonte = srcProdutos,

    Pipeline =
        fxStgPipelineGerar("srcProdutos"),

    Resultado =
        fxStgPadronizar(
            Fonte,
            Pipeline
        )
in
    Resultado;

shared tstStgAplicarTipos = let

    Fonte =
        srcVendas,

    Pipeline =
        fxStgPipelineGerar("tbVendas"),

    Resultado =
        fxStgAplicarTipos(
            Fonte,
            Pipeline
        )

in

    Resultado;

shared fxStgAplicar = (
    tabela as table,
    optional schema as nullable text
)
as table =>

let

    //----------------------------------------------------------------------
    // Configuração da execução
    //----------------------------------------------------------------------

    TemSchema =
        schema <> null,

    ProcessarTratamentos =
        TemSchema
        and
        fxParametro(
            "Processar_Tratamentos"
        ),

    ProcessarValidacoes =
        TemSchema
        and
        fxParametro(
            "Processar_Validacoes"
        ),

    //----------------------------------------------------------------------
    // Definição do Pipeline
    //----------------------------------------------------------------------

    Pipeline =
        if TemSchema then

            fxStgPipelineGerar(
                schema
            )

        else

            null,

    //----------------------------------------------------------------------
    // Preparação física da tabela
    //----------------------------------------------------------------------

    Preparada =
        fxStgPreparar(
            tabela
        ),

    //----------------------------------------------------------------------
    // Garantia da existência das colunas do Schema
    //----------------------------------------------------------------------

    ColunasGarantidas =
        if TemSchema then

            fxStgGarantirColunas(
                Preparada,
                Pipeline
            )

        else

            Preparada,

    //----------------------------------------------------------------------
    // Remoção das colunas que não pertencem ao Schema
    //----------------------------------------------------------------------

    ColunasRemovidas =
        if TemSchema then

            fxStgRemoverColunas(
                ColunasGarantidas,
                Pipeline
            )

        else

            ColunasGarantidas,

    //----------------------------------------------------------------------
    // Aplicação dos tratamentos definidos no Pipeline
    //----------------------------------------------------------------------

    Tratada =
        if ProcessarTratamentos then

            fxStgPadronizar(
                ColunasRemovidas,
                Pipeline
            )

        else

            ColunasRemovidas,

    //----------------------------------------------------------------------
    // Aplicação das validações definidas no Pipeline
    //----------------------------------------------------------------------

    Validada =
        if ProcessarValidacoes then

            fxStgAplicarValidacoes(
                Tratada,
                Pipeline,
                schema
            )

        else

            Tratada,

    //----------------------------------------------------------------------
    // Aplicação dos tipos definidos no Pipeline
    //----------------------------------------------------------------------

    Tipada =
        if TemSchema then

            fxStgAplicarTipos(
                Validada,
                Pipeline
            )

        else

            Validada,

    //----------------------------------------------------------------------
    // Ordenação final das colunas conforme o Schema
    //----------------------------------------------------------------------

    Ordenada =
        if TemSchema then

            fxStgOrdenarColunas(
                Tipada,
                Pipeline
            )

        else

            Tipada

in

    Ordenada;

shared tstStgAplicar = let

    Fonte = srcClientes,

    Resultado =
        fxStgAplicar(
            Fonte,
            "tbClientes"
            //null
        )
in
    Resultado;

shared fxStgAplicarValidacoes = /*
Responsabilidades:
- Executar as validações definidas no Pipeline.
- Nunca alterar os valores das colunas.
- Nunca converter tipos.
- Nunca adicionar ou remover linhas.
- Adicionar apenas a coluna "Ocorrencias".
*/

(
    tabela as table,
    pipeline as record,
    optional schema as nullable text
)
as table =>

let

    EstadoInicial =

        [

            NomesColunas =
                Table.ColumnNames(
                    tabela
                ),

            Colunas =
                Table.ToColumns(
                    tabela
                ),

            Ocorrencias =
                {}

        ],

    EstadoFinal =

        List.Accumulate(

            Record.FieldNames(
                pipeline
            ),

            EstadoInicial,

            (
                estado,
                coluna
            ) =>

                let

                    PipelineColuna =

                        Record.Field(
                            pipeline,
                            coluna
                        ),

                    Validacoes =
                        PipelineColuna[Validações] ?? {},

                    Posicao =

                        List.PositionOf(

                            estado[NomesColunas],

                            coluna

                        ),

                    Valores =

                        estado[Colunas]{Posicao},

                    Contexto =

                        [

                            Tabela =
                                schema,

                            Coluna =
                                coluna,

                            Definicao =
                                PipelineColuna

                        ],

                    OcorrenciasColuna =

                        if List.IsEmpty(
                            Validacoes
                        ) then

                            List.Repeat(

                                { null },

                                List.Count(
                                    Valores
                                )

                            )

                        else

                            fxStgAplicarValidacoesColuna(

                                Valores,

                                Validacoes,

                                Contexto

                            )

                in

                    estado &

                    [

                        Ocorrencias =

                            estado[Ocorrencias] &

                            {

                                OcorrenciasColuna

                            }

                    ]

        ),

    QuantidadeLinhas =

        if List.IsEmpty(
            EstadoFinal[Colunas]
        ) then

            0

        else

            List.Count(
                EstadoFinal[Colunas]{0}
            ),

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

                    if List.IsEmpty(
                        ListaLinha
                    ) then

                        null

                    else

                        ListaLinha

        ),

    Resultado =

        Table.FromColumns(

            EstadoFinal[Colunas] &

            {

                Ocorrencias

            },

            EstadoFinal[NomesColunas] &

            {

                "Ocorrencias"

            }

        )

in

    Resultado;

shared fxStgAplicarValidacoesColuna = /*
Função:
    fxStgAplicarValidacoesColuna

Responsabilidades:
- Executar todas as validações de uma única coluna.
- Executar todas as validações para todas as linhas.
- Acumular todas as ocorrências produzidas.
- Não interromper o processamento quando uma validação falha.

Não deve:
- Alterar valores.
- Construir tabelas.
- Mesclar ocorrências entre colunas.

Retorno:

{
    null,
    {
        Ocorrencia1,
        Ocorrencia2
    },
    null,
    {
        Ocorrencia3
    }
}
*/

(
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
                etapa
            ) =>

                let

                    Funcao =

                        etapa[Definição][Função],

                    Parametros =

                        Record.FieldOrDefault(

                            etapa,

                            "Parametros"

                        ),

                    Severidade =

                        Record.Field(

                            cfgParametrosSeveridades,

                            etapa[Definição][Severidade]

                        ),

                    ContextoEtapa =

                        contexto &

                        [

                            Etapa =
                                etapa,

                            Severidade =
                                Severidade

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

                                        ContextoEtapa

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
                cfgParametrosTipos,
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

    Fonte = srcProdutos,

    Pipeline =
        fxStgPipelineGerar("tbProdutos"),

    Resultado =
        fxStgAplicarValidacoes(
            Fonte,
            Pipeline,
            "Produtos"
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
    tabela as table,
    pipeline as record
)
as table =>

let

    ColunasSchema =

        Record.FieldNames(
            pipeline
        ),

    ColunasExistentes =

        Table.ColumnNames(
            tabela
        ),

    ColunasFaltantes =

        List.Difference(

            ColunasSchema,

            ColunasExistentes

        ),

    Resultado =

        if List.IsEmpty(ColunasFaltantes) then

            tabela

        else

            Table.Combine(

                {
                    tabela,
                    #table(
                        ColunasFaltantes,
                        {}
                    )
                }

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
    tabela as table,
    pipeline as record
)
as table =>

let

    Colunas =

        Record.FieldNames(
            pipeline
        ),

    Resultado =

        Table.SelectColumns(

            tabela,

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
    tabela as table,
    pipeline as record
)
as table =>

let

    Colunas =

        Record.FieldNames(

            pipeline

        ),

    Resultado =

        Table.ReorderColumns(

            tabela,

            Colunas,

            MissingField.Ignore

        )

in

    Resultado;