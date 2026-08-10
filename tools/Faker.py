import csv
import random
from faker import Faker  # type: ignore

fake = Faker("pt_BR")

NUM_ROWS = 100_000

headers = [
    "Nome",
    "CPF",
    "RG",
    "DataNascimento",
    "Sexo",
    "EstadoCivil",
    "Email",
    "Telefone",
    "Celular",
    "CEP",
    "Logradouro",
    "Numero",
    "Complemento",
    "Bairro",
    "Cidade",
    "UF",
    "Profissao",
    "RendaMensal",
    "DataCadastro",
    "ClienteAtivo"
]

ESTADOS_CIVIS = [
    "Solteiro",
    "Casado",
    "Divorciado",
    "Viúvo",
    "União Estável"
]

PROFISSOES = [
    "Administrador",
    "Advogado",
    "Analista de Sistemas",
    "Arquiteto",
    "Assistente Administrativo",
    "Atendente",
    "Autônomo",
    "Auxiliar Financeiro",
    "Comerciante",
    "Contador",
    "Dentista",
    "Desenvolvedor",
    "Designer",
    "Eletricista",
    "Empresário",
    "Engenheiro",
    "Enfermeiro",
    "Estudante",
    "Gerente",
    "Mecânico",
    "Médico",
    "Motorista",
    "Pedreiro",
    "Professor",
    "Programador",
    "Psicólogo",
    "Representante Comercial",
    "Técnico em Informática",
    "Vendedor"
]


def aplicar_defeito(tipo, registro):

    if tipo == "espacos":

        campo = random.choice([
            "Nome",
            "Logradouro",
            "Bairro",
            "Cidade",
            "Profissao"
        ])

        registro[campo] = random.choice([
            " " + registro[campo],
            registro[campo] + " ",
            "   " + registro[campo] + "   ",
            registro[campo].replace(" ", "  "),
            registro[campo] + "\t",
            registro[campo] + "\u00A0",
            registro[campo] + "\u200B"
        ])


    elif tipo == "cpf":

        cpf = "".join(filter(str.isdigit, registro["CPF"]))

        registro["CPF"] = random.choice([

            cpf,                                    # 12345678909

            f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}",

            f"{cpf[:3]} {cpf[3:6]} {cpf[6:9]} {cpf[9:]}",

            f"{cpf[:3]}-{cpf[3:6]}-{cpf[6:9]}-{cpf[9:]}",

            cpf[:-1],                              # 10 dígitos

            cpf + str(random.randint(0,9)),        # 12 dígitos

            cpf[:8],                               # curto

            "00000000000",

            "11111111111",

            "99999999999",

            "1234567890A",

            "ABC12345678",

            "",

            " ",

            "NULL",

            "N/A",

            "---"
        ])


    elif tipo == "rg":

        rg = "".join(filter(str.isalnum, registro["RG"]))

        registro["RG"] = random.choice([

            rg,

            f"{rg[:2]}.{rg[2:5]}.{rg[5:8]}-{rg[8:]}",

            rg.replace(".", ""),

            rg[:-1],

            rg + "X",

            "MG123456",

            "ISENTO",

            "",

            "NULL",

            "---",

            "123",

            "ABCDEFGHI"
        ])


    elif tipo == "cep":

        cep = "".join(filter(str.isdigit, registro["CEP"]))

        registro["CEP"] = random.choice([

            cep,

            f"{cep[:5]}-{cep[5:]}",

            f"{cep[:5]} {cep[5:]}",

            cep[:-1],

            cep + "9",

            "00000000",

            "99999-999",

            "ABC12345",

            "",

            "NULL",

            "-----"
        ])


    elif tipo == "telefone":

        tel = "".join(filter(str.isdigit, registro["Telefone"]))

        registro["Telefone"] = random.choice([

            tel,

            f"({tel[:2]}) {tel[2:7]}-{tel[7:]}",

            f"{tel[:2]} {tel[2:7]} {tel[7:]}",

            f"+55 ({tel[:2]}) {tel[2:7]}-{tel[7:]}",

            tel[:-2],

            tel + "99",

            "999999999",

            "abcdefgh",

            "",

            "NULL"
        ])


    elif tipo == "email":

        email = registro["Email"]

        registro["Email"] = random.choice([

            email.upper(),

            email.lower(),

            email.replace("@", ""),

            email.replace(".", ""),

            email.replace("@", "@@"),

            email.replace("@", "#"),

            email + " ",

            "usuario@",

            "@gmail.com",

            "usuario.gmail.com",

            "",

            "NULL"
        ])


    elif tipo == "data":

        registro["DataNascimento"] = random.choice([

            "31/02/2024",

            "99/99/9999",

            "2024-15-99",

            "15-08-1990",

            "1990/08/15",

            "19900815",

            "",

            "NULL",

            "abc"
        ])


    elif tipo == "valor":

        valor = registro["RendaMensal"]

        registro["RendaMensal"] = random.choice([

            valor,

            str(valor).replace(".", ","),

            f"R$ {valor:,.2f}",

            f"R${valor}",

            "1.234,56",

            "1234.56",

            "N/A",

            "",

            "NULL",

            "---"
        ])


    elif tipo == "nulos":

        campo = random.choice(list(registro.keys()))

        registro[campo] = random.choice([

            "",

            " ",

            "   ",

            "NULL",

            "null",

            "N/A",

            "-",

            "---",

            "Não Informado",

            "Sem Informação"
        ])


    elif tipo == "especial":

        campo = random.choice([
            "Nome",
            "Logradouro",
            "Bairro",
            "Cidade",
            "Profissao"
        ])

        registro[campo] += random.choice([

            "@#$%",

            "***",

            "///",

            "\\\\\\",

            "<>",

            "{}[]",

            "😊",

            "©®™",

            "\t"
        ])


    elif tipo == "duplicado":

        registro["CPF"] = "111.111.111-11"
        registro["Email"] = "cliente@empresa.com"


    return registro

with open(
    "clientes.csv",
    "w",
    newline="",
    encoding="utf-8-sig"
) as csvfile:

    writer = csv.DictWriter(csvfile, fieldnames=headers, delimiter=";")
    writer.writeheader()

    for _ in range(NUM_ROWS):

        sexo = random.choice(["Masculino", "Feminino"])

        registro = {

            "Nome":
                fake.name_male()
                if sexo == "Masculino"
                else fake.name_female(),

            "CPF":
                fake.cpf(),

            "RG":
                fake.rg(),

            "DataNascimento":
                fake.date_of_birth(
                    minimum_age=18,
                    maximum_age=90
                ).strftime("%d/%m/%Y"),

            "Sexo":
                sexo,

            "EstadoCivil":
                random.choice(ESTADOS_CIVIS),

            "Email":
                fake.email(),

            "Telefone":
                fake.phone_number(),

            "Celular":
                fake.cellphone_number(),

            "CEP":
                fake.postcode(),

            "Logradouro":
                fake.street_name(),

            "Numero":
                random.randint(1,9999),

            "Complemento":
                random.choice([
                    "",
                    "Apto 101",
                    "Apto 302",
                    "Casa",
                    "Fundos",
                    "Bloco A",
                    "Bloco B",
                    "Sala 201",
                    "Loja 05"
                ]),

            "Bairro":
                fake.bairro(),

            "Cidade":
                fake.city(),

            "UF":
                fake.estado_sigla(),

            "Profissao":
                random.choice(PROFISSOES),

            "RendaMensal":
                round(random.uniform(1500,50000),2),

            "DataCadastro":
                fake.date_between(
                    start_date="-10y",
                    end_date="today"
                ).strftime("%d/%m/%Y"),

            "ClienteAtivo":
                random.choice(["Sim","Não"])
        }

        prob = random.random()

        defeitos = []

        if random.random() < 0.05:
            defeitos.append("espacos")

        if random.random() < 0.03:
            defeitos.append("cpf")

        if random.random() < 0.02:
            defeitos.append("rg")

        if random.random() < 0.02:
            defeitos.append("cep")

        if random.random() < 0.03:
            defeitos.append("telefone")

        if random.random() < 0.02:
            defeitos.append("email")

        if random.random() < 0.02:
            defeitos.append("data")

        if random.random() < 0.02:
            defeitos.append("valor")

        if random.random() < 0.03:
            defeitos.append("nulos")

        if random.random() < 0.01:
            defeitos.append("especial")

        for defeito in defeitos:
            registro = aplicar_defeito(defeito, registro)

        writer.writerow(registro)