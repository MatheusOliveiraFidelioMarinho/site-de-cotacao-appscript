# Sistema de Compras A+.com

Site para cotar, aprovar e acompanhar compras. O banco de dados é uma planilha do Google, e o site é um App da Web do Google Apps Script.

## Como funciona

Tudo acontece numa tela só: cada **projeto** é um bloco com os seus itens e os totais (em cotação, aprovado e comprado). Clique no nome do projeto para recolher ou abrir o bloco.

1. **+ Novo projeto** → nome e cliente.
2. **+ Itens** no projeto → lance vários itens de uma vez (descrição, quantidade, unidade). Dá para colar direto do Excel.
3. **+ Cotar** na linha do item → fornecedor, onde foi cotado, link, valor, frete, pagamento e prazo. Use "Salvar e adicionar outra" para cadastrar várias cotações seguidas.
4. **Enviar p/ aprovação**: pela linha, ou marque vários itens e envie todos juntos.
5. **Aprovador**: na própria linha escolhe a cotação (a mais barata já vem selecionada) e clica **✓ Aprovar**. Também pode marcar vários itens e usar **Aprovar marcados**.
6. Depois: **🛒 Comprado** e **📦 Recebido**, também com um clique na linha.

Clicar no item abre o detalhe ao lado, com as cotações comparadas em cartões, a **conversa** entre comprador e aprovador, o histórico e as opções de **devolver para cotação**, **reprovar** (com motivo), editar e cancelar. Mensagens não lidas aparecem com um ponto azul.

O aprovador também pode cotar e lançar itens e pode **encerrar** um projeto terminado, que então some da tela (marque "Encerrados" para ver de novo). A tela se atualiza sozinha a cada 45 segundos.

## Arquivos

| Pasta / arquivo | Para quê |
|---|---|
| `apps-script/Code.gs` | Backend: lê e grava na planilha e aplica as regras |
| `apps-script/Index.html`, `Estilos.html`, `App.html` | O site |
| `apps-script/Logo.html` | Logo da empresa embutida |
| `apps-script/appsscript.json` | Configuração do projeto (fuso, permissões do app) |
| `ferramentas/` | Prévia local para testes. **Não vai para o Google.** |

## Instalação (uma vez, uns 10 minutos)

1. No Google Drive da empresa, crie uma **planilha nova** (ex.: "Compras A+.com").
2. Na planilha, vá em **Extensões > Apps Script**.
3. No editor:
   - Apague o conteúdo do `Código.gs` e cole o conteúdo de `apps-script/Code.gs`.
   - Clique em **+ > HTML** e crie 4 arquivos com estes nomes exatos (sem ".html"): `Index`, `Estilos`, `App` e `Logo`. Cole em cada um o conteúdo do arquivo correspondente. O `Logo.html` é uma linha só e bem comprida; copie inteira.
   - Clique na engrenagem (**Configurações do projeto**), marque **"Mostrar o arquivo de manifesto appsscript.json no editor"**, volte ao editor e substitua o conteúdo de `appsscript.json` pelo arquivo daqui.
   - Salve (Ctrl+S).
4. Volte à planilha e recarregue a página. Vai aparecer o menu **Sistema de Compras > Configurar planilha**. Clique nele e autorize. Isso cria as abas e cadastra você como APROVADOR.
5. Na aba **Usuarios**, cadastre as pessoas:

   | Email | Nome | Papel | Ativo |
   |---|---|---|---|
   | fulano@... | Fulano | COMPRADOR | SIM |
   | ciclano@... | Ciclano | COMPRADOR | SIM |
   | aprovador@... | Aprovador | APROVADOR | SIM |

6. **Compartilhe a planilha** com essas pessoas como **Editor**. É necessário porque o site roda com a conta Google de cada um.
7. No Apps Script: **Implantar > Nova implantação**, tipo **App da Web**:
   - Executar como: **Usuário que acessa o app da Web**
   - Quem pode acessar: **Qualquer pessoa com Conta do Google**
   - Clique em **Implantar** e copie a URL. Esse é o link do sistema, para mandar ao pessoal e salvar nos favoritos.

No primeiro acesso, cada pessoa precisa clicar em "Autorizar" uma vez.

### Se todos usam e-mail da empresa (Google Workspace, mesmo domínio)

Dá para não compartilhar a planilha com ninguém. Na implantação, use **Executar como: Eu** e **Quem pode acessar: Qualquer pessoa em [seu domínio]**, e troque no `appsscript.json` para:

```json
"webapp": { "executeAs": "USER_DEPLOYING", "access": "DOMAIN" }
```

Assim só o site mexe na planilha. **Não use** esse modo com contas @gmail.com, porque o Google não informa o e-mail de quem acessa e o sistema não consegue identificar a pessoa.

## Atualizar o sistema depois de mudanças

Cole os arquivos alterados no editor e vá em **Implantar > Gerenciar implantações > ✏️ (editar) > Versão: Nova versão > Implantar**. A URL continua a mesma.

## Prévia local (para testar mudanças)

```
node ferramentas/preview.js
```

Abra http://localhost:5180. Ela roda o mesmo `Code.gs` com uma planilha simulada no navegador e já vem com dados de exemplo. A barra amarela no canto troca de usuário (aprovadora, compradores e alguém sem acesso) e permite resetar os dados.

## Abas da planilha

- **Usuarios**: quem acessa e com qual papel. Para tirar o acesso de alguém, troque Ativo para `NAO`.
- **Projetos**: criados pelo próprio site ("+ Novo projeto…"). Para esconder um projeto da lista, troque Ativo para `NAO`.
- **Itens**: as solicitações e o status de cada uma.
- **Cotacoes**: as cotações de cada item. As removidas ficam com Ativa = `NAO`.
- **Comentarios** e **Historico**: a conversa e o registro de ações.

Evite editar as abas Itens e Cotacoes à mão, para não quebrar o fluxo. Para relatórios, crie abas novas com fórmulas ou tabelas dinâmicas lendo essas abas.
