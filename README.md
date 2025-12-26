# Sistema de Controle de Acesso e Proteção de Abas no Excel (VBA)

Este repositório contém um projeto em **VBA para Excel** que implementa:

- Tela de **Login** e **Cadastro** de usuários (UserForm `frmAcesso`).
- Controle de **perfis de acesso** (`admin` e `usuario`).
- **Proteção/Exibição** dinâmica de abas (worksheets) conforme perfil.
- **Histórico de logins** e registro do **último acesso** por usuário.
- Botões de ação para **consulta** (somente visualizar), **edição** (desbloquear para editar) e **sair**.
- Aplicação de **AutoFilter** nas principais abas para facilitar consulta.
- Mecanismo de **bloqueio temporário** após tentativas de login falhas.

---

## Estrutura esperada do arquivo Excel

O projeto pressupõe a existência das seguintes **abas** (sheets) no workbook:

- `Usuarios`: cadastro sem senha (colunas: **A** Usuário, **B** Tipo, **C** Último Acesso). A coluna **B** pode ser protegida/ocultada conforme regras do sistema.
- `SenhasUsuarios`: armazena credenciais (colunas: **A** Usuário, **B** Senha, **C** Tipo). **Atenção**: as senhas aqui estão **em texto claro** na versão final; veja recomendações de segurança abaixo.
- `HistoricoLogins`: log de entradas (colunas: **A** Usuário, **B** Data, **C** Hora, **D** Tipo, **E** Ação).
- `HistoricoEdicoes`: opcional; quando presente, tem sua proteção ajustada pelo sistema.
- `Base de dados`: dados do negócio; sempre **desprotegida** ao entrar.
- `Usuarios` e demais abas adicionais do seu caso de uso.

> O nome das abas é **sensível**: o código referencia explicitamente cada uma.

---

## Componentes principais

### 1) Eventos do Workbook
- `Workbook_Open`: valida a existência da aba `Usuarios`, oculta todas as demais com `xlSheetVeryHidden`, oculta a coluna **B** (senhas num design anterior), protege de volta, e **exibe o formulário de login** (`frmAcesso`).
- `Workbook_BeforeClose`: salva automaticamente o arquivo ao fechar.

### 2) Formulário `frmAcesso`
Inclui **frames** de Login (`fraLogin`) e Cadastro (`fraCadastro`), além de botões:

- `btnEntrar_Click`: realiza autenticação consultando a aba `SenhasUsuarios`; se ok:
  - Mostra as abas (`MostrarAbas`).
  - Atualiza **Último Acesso** em `Usuarios` (formato `dd/mm/yyyy hh:mm:ss`).
  - Registra no `HistoricoLogins` (usuário, data, hora, tipo, ação "Login").
  - Controla proteção conforme **tipo de usuário**:
    - **admin**: desbloqueia e **torna visível** `SenhasUsuarios`; desprotege `Usuarios`, `HistoricoEdicoes`, `HistoricoLogins`.
    - **usuario**: protege `Usuarios`, `HistoricoEdicoes`, `HistoricoLogins` e oculta `SenhasUsuarios` com `xlSheetVeryHidden`.
  - Ativa **AutoFilter** nas abas principais.
  - Reseta contador de **tentativas**.
- Se a autenticação falhar, incrementa `tentativas`; após **3 erros**, desabilita o botão de entrar por **5 minutos** usando `Application.OnTime` e chama `ReativarLogin` depois.

- `btnCadastrar_Click`: cadastra novo usuário. Na versão atual:
  - Escreve **usuário e tipo** em `Usuarios`.
  - Escreve **usuário, senha (texto claro) e tipo** em `SenhasUsuarios`.
  - Define **tipo** automaticamente: primeiras linhas (até a 3ª) viram `admin`; demais, `usuario`.

- `btnConsulta_Click` / `btnConsultar_Click`: torna todas as abas **visíveis** e **protege** para *visualização* (UserInterfaceOnly), útil para navegação segura.

- `btnEditarMaterial_Click`: torna todas as abas **visíveis** e **desprotege** para permitir **edição**.

- `btnSair_Click`: fecha o formulário e informa que o usuário saiu.

- `UserForm_Initialize`: prepara estado inicial (mostra login, oculta cadastro, zera tentativas).

### 3) Módulos auxiliares
- `MostrarAbas`: torna visíveis todas as abas exceto `Usuarios` quando logado.
- `DesprotegerTodasAsAbas`: utilitário para deixar todas as abas visíveis.
- `ReativarLogin`: reabilita o botão de login após o tempo de punição.
- `CriptografarSenha`: função simples (Caesar shift +3) usada no **design anterior**; **não é** aplicada na versão com `SenhasUsuarios` (lá a senha está em **texto claro**).

---

## Fluxo de uso

1. **Abrir o arquivo**: o evento `Workbook_Open` oculta abas e exibe `frmAcesso`.
2. **Login**: informe *Usuário* e *Senha* cadastrados em `SenhasUsuarios`.
3. **Pós-login**: dependendo do **tipo** (`admin`/`usuario`), o sistema ajusta proteção e visibilidade.
4. **Consulta/Edição**: use os botões do menu para liberar somente visualização ou permitir edição.
5. **Cadastro**: via `fraCadastro`, crie novos usuários; o tipo é determinado automaticamente pela posição.
6. **Saída**: use `btnSair_Click`; o arquivo é salvo ao fechar.

---

## Instalação e configuração

1. **Habilite macros** no Excel (Arquivo → Opções → Central de Confiabilidade → Configurações de Macro).
2. **Insira as abas** com os nomes exatamente como descritos.
3. **Importe o UserForm** `frmAcesso` e os módulos com o código deste repositório.
4. **Ajuste a senha de proteção** (string `"visualizar"`) se desejar alterar o segredo usado em `Protect/Unprotect`.
5. Opcional: configure formatação e campos adicionais conforme seu negócio.

---

## Boas práticas e recomendações de segurança

- **Senhas em texto claro**: a aba `SenhasUsuarios` guarda senhas **sem hash**. Recomenda-se:
  - Substituir por **hash seguro** (ex.: SHA-256 via biblioteca externa ou implementação própria) e comparar o hash.
  - Se quiser manter algo simples em VBA, *no mínimo* use uma função de hash (não reversível). O `CriptografarSenha` atual é apenas um deslocamento de caracteres (**não seguro**).
- **Proteção de planilhas** em Excel não é equivalente a criptografia: usuários avançados podem burlar. Use proteção como **camada de conveniência**, não de segurança absoluta.
- **Ocultação `xlSheetVeryHidden`** dificulta acesso casual, mas não impede acesso via VBA.
- **Controle de perfis**: centralize a lista de permissões em uma aba/estrutura única para evitar divergências.

---

## Personalização

- Troque os nomes das abas e os rótulos do formulário conforme seu domínio (ex.: Banco de Germoplasma, Materiais, etc.).
- Adicione novos perfis (ex.: `consulta`, `editor`) e ajuste a lógica de proteção.
- Expanda o registro de histórico para incluir **ações de edição** na `HistoricoEdicoes`.

---

## Limitações conhecidas

- Dependência de **Excel Desktop** com macros habilitadas (Windows/Mac).
- Segurança **limitada**: não há criptografia de arquivo nem armazenamento seguro de senhas.
- Alguns trechos usam `On Error Resume Next`; considerar tratamento de erros mais robusto.

