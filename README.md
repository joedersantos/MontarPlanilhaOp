
# MontarPlanilha

## Descrição do Projeto
Aplicação protótipo CLI que auxilia no processamento de um csv para uma planilha Excel.

## Decisões técnicas do projeto:
Esse projeto foi desenvolvendo utilizando .NET 5.0 com a linguagem C# e a biblioteca OpenXML. 


## Pré-requesitos

:warning: [.NET 5.0](https://dotnet.microsoft.com/download)


## Como instalar a aplicação :arrow_forward:

:heavy_check_mark:No terminal, entre na pasta raiz do projeto:

```
> MontarPlanilha
```

:heavy_check_mark:Instale a ferramenta globalmente usando este comando:

```
dotnet tool install --global --add-source src MontaPlanilhaOp
```

## Como a aplicação deve funcionar:

:trophy: O programa espera receber como entrada de um arquivo contendo linhas no formato csv na entrada padrão (stdin) e fornecera uma saída (stdout) indicando o andamento do processamento.
:trophy: Ao finalizar o processamento um arquivo .xlsx será disponibilizado na pasta onde o programa foi processado.


## Executar a aplicação:

:trophy: Navegue até o diretório onde se encontra o arquivo de input a ser passado como parâmetro e execute o comando:

```
MontaPlanilhaOp < [nome_do_arquivo_sem_extenção>] 
```
> **_Nota:_** Afim de exemplificar e fazer um teste foi incluído um arquivo [csvTeste] na raiz do projeto. Execute na pasta raiz:
```
MontaPlanilhaOp < csvTeste
```


## Caso queira desinstalar a ferramenta globalmente usando este comando: 
```
dotnet tool uninstall --global authorize
```
