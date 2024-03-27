# POC utilizando a biblioteca Exceljs

Biblioteca Exceljs: https://github.com/exceljs/exceljs



## Links

#### Cores ARGB
##### Converter cores Hexadecimais para ARGB
-  https://www.myfixguide.com/color-converter/
__________________________________

| POSSIVEL | ESTUDO |
| --- | --- |
| ✔️ | Customização das linhas e colunas: fonte, cor, tabulação |
| ✔️ | Impressão de imagem: queremos gerar com logos no cabeçalho |
| ❌ | Impressão de gráficos(https://github.com/exceljs/exceljs/issues/1173, https://github.com/exceljs/exceljs/issues/307) |

Observação de estudo:
>
> Foi estudo a possibilidade de lidar com planilhas como template, ou seja, editar uma planilha existente.
>
> Nesse estudo foi encontrado um problema, quando se edita uma planilha template e a gera novamente para o usuario.
>
> A planilha gera erros de formatação ao tentar abrir, deixando de ser usual. 
>
> Um desenvolvedor que utiliza a biblioteca, abriu um PR para corrigir o bug, mas segue aberto(https://github.com/exceljs/exceljs/pull/2185/files/2f1cd6e54bb84bd78aca6834ab5469c3c3e383f2).

