# SenderApp for WhatsApp



##### Server Specifications:  
![](https://img.shields.io/static/v1?label=Hosted%20on&message=Google%20Apps%20Scripts%20%2F%20Google%20Cloud%20Project&color=important)

##### Client Specifications:
![](https://img.shields.io/static/v1?label=SenderApp&message=v1.1&color=blue) ![](https://img.shields.io/static/v1?label=Python&message=v.3.7&color=brightgreen)


*SenderApp for WhatsApp* é um bot de envio de mensagem em lote que possibilita o envio de até **100 mensagens de texto** em poucos segundos, **a cada hora**. 

> O uso indevido desta ferramenta pode ocasionar no banimento da sua conta do WhatsApp. Recomenda-se leitura do Artigo Técnico [How WhatsApp Fights Bulk Messaging and Automated Behavior](https://scontent.fgpb4-1.fna.fbcdn.net/v/t39.8562-6/299842918_397263792546125_6219151513993243581_n.pdf?_nc_cat=107&ccb=1-7&_nc_sid=ae5e01&_nc_ohc=Rmmm-0GP-bIAX_2WITo&_nc_ht=scontent.fgpb4-1.fna&oh=00_AfCAG5ZpQQEYz8hy2L-Ca2bHU8bU3jwmRYomsLcgehQ8og&oe=636A5C8A).


## Como Utilizar

Uma versão do serviços, hospedado na Google (WebApps) é utilizado como servidor de comunicação da aplicação. 

Para utilizar o serviço, basta [criar sua conta](https://script.google.com/macros/s/AKfycbyyVrXZ2nmgwuPBcrrL2OWQWVbLKf_PkVWNIXT_kZ4UAgkhk0HrGxm7MgvxVtMx9PePjg/exec?a=r) e seguir as orientações recomendadas.
**DICA**: Caso queira mais informações sobre o produto ou de como realizar a instalação do mesmo, recomenda-se a leitura do sessão [ajuda](https://script.google.com/macros/s/AKfycbyyVrXZ2nmgwuPBcrrL2OWQWVbLKf_PkVWNIXT_kZ4UAgkhk0HrGxm7MgvxVtMx9PePjg/exec?a=h).

---

> **IMPORTANTE**: Caso deseje implementar seu próprio servidor ao invés de utilizar a versão citada acima, siga as orientações disponíveis em  './scripts/more-info.md'

## Proximos Passos - Planilha

Ao finalizar seu cadastro, uma planilha google será compartilhada com seu endereço de email.

#### Configurações da Planilha
O menu de configurações da planilha esta oculta por padrão. Para acessá-lo, clique no botão menu, nas abas da planilha e seleciona a planilha "Configurações". Todas as configurações possuem notas com orientações.

#### Dados Variáveis
O SenderApp permite a utilização de `02 Dados Variáveis`, sendo `@Nome` e `@Info`.
**Dica**: Deve-se respeitar a letra maiúscula da variável para que a mesma seja substituida corretamente. 

#### Configurando Mensagem
A mensagem a ser enviada deve ser informada em `'Mensagens!A3'` da planilha.
A mensagem pode conter um tamanho máximo de caracteres, e caso exceda este limite, a célula ficará vermelha e o bot não poderá ser iniciado.

A mensagem são apenas em formato `TEXTO` e  podem conter `links` e `formatação de texto`.
**Dica**: Para quebrar linha no corpo da mensagem, utilize o atalho `'ctrl' + 'enter'`.

---

**IMPORTANTE**: Para que tudo funcione corretamente, todos os dados das planilhas **Mensagem** e **Contatos**, devem ser formatados como `Texto simples`. Ao colar informações de outras fontes de dados, verifique se os dados estão formatados em texto simples. 


 ## Próximos Passos - SenderApp Desktop

Após fazer o [download do aplicativo](https://mega.nz/file/xcpnkZqb#stGI2EibhJ7b3rmou6reKFLXrxI5cpuy6v6alyDV8kU), não se esqueça de **executa-lo como administrador.**

Finalize a configuração informando seu `email` e seu `token de ativação`. 
**Dica**: Seu token de ativação é enviado para seu email no momento em que você finaliza seu cadastro.

## Próximos Passos - Primeiro envio

É necessário possui o WhatsApp Desktop instalado e aberto em seu computador. 
Não se preocupe, o primeiro envio será enviado para um número teste preparado para receber estas mensagens. 

Com o WhatsApp logado, pronto para ser utilizado, e o SenderApp sendo executado simultaneamente, basta clicar em  **`Iniar Envio`**. Lembre-se, é necessário executar como administrador para que o bot consiga enviar de fato as mensagens.

Ao finalizar o envio, uma mensagem de sucesso será exibida e você terá utilizado `10 cotas`.

#  Cotas de Mensagens
Como uma forma de limitar o uso indevido da ferramenta, um sistemas de cotas de envio é estabelecido, onde as regras são
- Cada usuário possui `100 cotas por padrão`.
- **Cada mensagem enviada, onde a requisição de envio é realizada com sucesso, independente da resposta do solicitação, consome** `01 cota`.
- Para o bot ser iniciado é necessário possui ao menos `10 números válidos`em sua lista de contatos.
- O Servidor *SenderApp for WhatsApp* recicla `02 cotas a cada minuto`.

Portanto, não é necessário esperar 01 hora obrigatoriamente para realizar o próximo envio, basta ter no mínimo 10 cotas disponíveis e contatos válidos para realizar envios com mais frequência. 

# Roadmap - Futuras Implementações:

 - [ ] Envio de Imagens
 - [ ] Envio de Audios
 - [ ] Envios Agendados / Programados - com Notificações
 

