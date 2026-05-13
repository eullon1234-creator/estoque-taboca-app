# Documentação técnica — Controle de Estoque (UHE Estrela / Taboca)

Este documento descreve a arquitetura, dados, fluxos e operação do aplicativo web de almoxarifado contido neste repositório. Versão de referência do app no código: **EULLON 2.0.4** (rodapé da sidebar e `manifest.json`); o `package.json` pode mostrar outro número de versão — use o manifest e o cache do service worker como referência de release.

---

## 1. Visão geral

### 1.1 Propósito

Sistema para **cadastro de itens**, **entradas** (com NF e fornecedor), **saídas rastreadas** (principalmente via **requisições**), **requisições de material** com impressão/PDF, **histórico e auditoria**, **locais**, **cautela de ferramentas**, **baixa de RM** (marcação administrativa no histórico), **sugestão de compras**, **relatórios**, **plaquinhas**, módulo adicional **Estoque Estrela** (coleções dedicadas), **PWA** (instalação e cache) e **backup/restauração** em JSON.

### 1.2 Stack tecnológica

| Camada | Tecnologia |
|--------|------------|
| Interface | HTML5, Tailwind CSS (build local), `assets/css/main.css` |
| Lógica | Um único módulo ES: `assets/js/main.js` (import maps via URLs do Firebase CDN) |
| Backend / dados | **Google Firebase** — **Cloud Firestore** (sem Firebase Authentication no login; ver §4) |
| Gráficos | Chart.js (CDN) |
| Planilhas | xlsx-js-style (CDN) |
| PDF / captura | jsPDF + html2canvas (CDN) |
| Offline / instalável | `manifest.json` + `sw.js` (Service Worker) |

### 1.3 Estrutura de pastas relevante

```
estoque-taboca-app-1/
├── index.html              # Shell da SPA: login, layout, todas as “views”
├── manifest.json           # PWA (nome, ícones, tema)
├── sw.js                   # Service Worker (cache versionado)
├── assets/
│   ├── js/main.js          # Toda a lógica do app (~7k linhas)
│   ├── css/
│   │   ├── tailwind.input.css
│   │   ├── tailwind.css    # Gerado pelo build Tailwind
│   │   └── main.css        # Estilos complementares
│   └── icons/              # icon-192.png, icon-512.png
├── tailwind.config.js
├── package.json            # Scripts npm: build/watch do CSS
└── README.md               # Visão geral curta + link GitHub Pages
```

Não há framework (React/Vue); a navegação entre telas é feita por **mostrar/ocultar** elementos com classe `view-content` e `hidden`.

---

## 2. Execução e build

### 2.1 Rodar localmente

1. Abra a pasta do projeto no editor.
2. Sirva os arquivos por **HTTP** (extensão Live Server, `npx serve`, IIS, etc.). Abrir `index.html` direto pelo protocolo `file://` pode quebrar ES modules ou service worker.
3. Opcional: recompilar CSS após alterar classes no HTML/JS:
   - `npm install`
   - `npm run build:css` — gera `assets/css/tailwind.css` minificado
   - `npm run watch:css` — modo watch durante desenvolvimento

### 2.2 Publicação (GitHub Pages)

Conforme `README.md`: configurar Pages na branch principal, pasta raiz. Lembre-se de que o **Firestore** continua sendo o backend na nuvem; o Pages só hospeda o front-end.

---

## 3. Arquitetura do front-end

### 3.1 Fluxo de carregamento

1. `index.html` carrega CDNs (Chart.js, XLSX, jsPDF, html2canvas), depois `tailwind.css`, `main.css` e `main.js` como **módulo**.
2. Overlay de loader (`#loader-overlay`) até dados mínimos estarem prontos.
3. Se existir `localStorage.appUser`, a sessão é restaurada e `initializeAppSession` roda sem passar pela tela de login.
4. Caso contrário, exibe `#login-screen`.
5. Após login, `setupListeners` → `syncRealtimeListeners` → listeners `onSnapshot` do Firestore preenchem arrays em memória (`products`, `history`, etc.) e disparam renders.

### 3.2 Navegação (views)

A função `switchView(viewId)`:

- Define `currentViewId`, alterna visibilidade dos blocos `.view-content`, marca abas ativas (sidebar desktop + bottom nav mobile).
- Dispara renders específicos (ex.: `renderEntryView`, `updateDashboard`, `renderEstrelaView`).
- Chama `syncRealtimeListeners()` para **ligar/desligar** listeners do módulo Estrela conforme a aba (economia de leituras quando a aba Estrela não está visível).

Principais IDs de view (atributo `data-view` / `id`): `dashboard-view`, `inventory-view`, `add-product-view`, `entry-view`, `plaques-view`, `requisitions-view`, `exit-log-view`, `activity-log-view`, `rm-view`, `tool-loans-view`, `estrela-view`, `compras-view`, `reports-view`.

### 3.3 Layout responsivo

- **Desktop (lg+):** sidebar fixa com seções (Geral, Estoque, Operações, UHE Estrela, Análise, Sistema).
- **Mobile:** barra inferior + drawer “Mais” para o restante das rotas.

### 3.4 PWA e atualização

- **`manifest.json`:** nome curto “Estrela”, `display: standalone`, ícones 192/512.
- **`sw.js`:** `CACHE_NAME` versionado (`eullon-v2.0.4`). Em `fetch`, navegação e URLs com query `?v=` ou `?atualizar=` priorizam **rede** para pegar HTML/JS/CSS novos; demais GETs podem servir do cache.
- Botão **Atualizar** no header força sincronização/limpeza de cache (comportamento ligado ao SW no `main.js`).
- Botão **Instalar** aparece quando o navegador expõe `beforeinstallprompt`.

---

## 4. Autenticação e sessão

### 4.1 Modelo

O app **não usa Firebase Auth** para o formulário de login. A autenticação é **custom**:

1. Usuário e senha digitados; senha é hasheada com **SHA-256** (`crypto.subtle`) no cliente.
2. **Usuários embutidos (built-in):** dois IDs fixos no código (`BUILTIN_USERS` em `main.js`) com papel `admin` e `obraId` distintos (`uhe_estrela` e `pch_taboca`). A senha é comparada ao hash da senha configurada no código-fonte.
3. **Demais usuários:** documento em Firestore em `.../users/{userId}` com campo `passwordHash` igual ao hash informado. Campos típicos: `displayName`, `role`, etc.

O objeto de sessão salvo em `localStorage` sob a chave **`appUser`** contém pelo menos: `uid`, `displayName`, `role`, `obraId`.

### 4.2 Papéis e permissões

Definidos em `PERMISSIONS`:

| Papel | Permissões |
|-------|------------|
| **admin** | create, read, update, delete, export, import, manage_users, settings |
| **operador** | create, read, update, export |
| **visualizador** | read, export |

- Usuário novo encontrado em `users` sem documento prévio pode ser criado como **visualizador** (lógica em `getUserRole` quando usada em outros fluxos).
- Botões de criar/editar/importar/backup/configurações respeitam `hasPermission`.
- Exclusões sensíveis no estoque usam modal com **PIN** (`STOCK_DELETE_PIN`, valor fixo no código — proteção contra clique acidental, **não** substitui controles de segurança fortes).

### 4.3 Avisos de segurança (importante)

- Credenciais built-in e PIN estão **no código-fonte** visível a qualquer um com acesso ao repositório ou ao JS publicado. Para ambiente sério, use **Firebase Auth**, **regras Firestore** restritivas e **remoção de segredos** do bundle.
- `passwordHash` no Firestore com hash só no cliente ainda é vulnerável a enumeração offline se as regras permitirem leitura ampla.
- A segurança real depende das **Security Rules** do projeto Firebase (não versionadas neste repo).

---

## 5. Firebase e Firestore

### 5.1 Configuração

Objeto `firebaseConfig` e `appId` estão em `assets/js/main.js`:

- `appId = 'ferramentaria-estoque'` — prefixo lógico dos caminhos.
- Inicialização do app + Firestore com **cache persistente** (`persistentLocalCache` + `persistentMultipleTabManager`) quando suportado; fallback para `getFirestore`.

### 5.2 Caminhos das coleções (multi-obra)

Base por obra (`initializeAppSession`):

- Se `obraId === 'uhe_estrela'` (padrão):  
  `obraBase = /artifacts/{appId}/public/data`
- Caso contrário (ex.: `pch_taboca`):  
  `obraBase = /artifacts/{appId}/public/data/obras/{obraId}`

Sob `obraBase`:

| Coleção / doc | Uso |
|---------------|-----|
| `products` | Itens de estoque principal |
| `history` | Movimentações / saídas / entradas registradas no histórico unificado |
| `requisitions` | Requisições de material |
| `tool_loans` | Cautelas (empréstimos/devoluções) |
| `locations` | Cadastro de locais |
| `app_settings/main` | Documento único: `appName`, `logoUrl`, etc. |

**Sempre na raiz `public/data` (não por obra):**

| Caminho | Uso |
|---------|-----|
| `users` | Perfis, hash de senha, papéis |
| `suggestions` | Feedback/sugestões enviadas pelo app |
| `estrela_products`, `estrela_entries`, `estrela_exits` | Módulo **Estoque Estrela** (paralelo ao estoque principal) |

A navegação “Estoque Estrela” só é exibida quando `obraId === 'uhe_estrela'`.

### 5.3 Tempo real

`onSnapshot` mantém `products`, `history`, `requisitions`, `toolLoans`, `locations` e `app_settings` atualizados. O carregamento inicial considera `locations` como gatilho para marcar `isDataLoaded` e esconder o skeleton (ver código para detalhe de ordem).

Comentário no código: listeners **não** são removidos ao mudar foco da aba (evita releitura cara ao voltar); **logout** encerra listeners.

### 5.4 Histórico (`history`)

Campos relevantes em `addHistoryEntry`: `productId`, códigos/nome do produto, `type`, `quantity`, `newTotal`, `withdrawnBy`, `teamLeader`, `applicationLocation`, `obra`, `details`, **`rmProcessed`** (boolean, controle da tela Baixa RM), `date`/`timestamp` com `serverTimestamp()`, `performedBy` (auditoria).

---

## 6. Domínio funcional (telas)

### 6.1 Painel (`dashboard-view`)

Indicadores agregados: totais de produtos/unidades, estoque baixo, requisições pendentes, gráficos (Chart.js), atividade recente, cautelas do dia, etc. Função principal: `updateDashboard`.

### 6.2 Estoque (`inventory-view`)

- Lista paginada (**50 itens por página**), busca, filtros (todos / estoque baixo), ordenação (nome, código RM, local), filtro por local.
- Ações por item: editar, ajustar quantidade, excluir (conforme permissão + PIN onde aplicável).
- Seleção múltipla: excluir em lote, **iniciar requisição** com itens selecionados.
- Exportação (Excel/CSV conforme implementação no botão export).

### 6.3 Cadastrar (`add-product-view`)

Formulário: nome, foto por URL opcional, código RM opcional, código interno (gerado), localização, grupo, unidade, quantidade inicial, mínimo.

- **Importação CSV** pelo botão na mesma área.
- Grupos válidos alinhados a `VALID_PRODUCT_GROUPS`; existe **inferência automática** de grupo para produtos com grupo vazio, “Outros” genérico ou “Ferramentas” (`autoFillMissingProductGroups`) para usuários com `update`.

### 6.4 Registrar entrada (`entry-view`)

Busca de produto, quantidade, NF obrigatória, fornecedor, observação. Atualiza estoque e grava histórico.

### 6.5 Saídas

A view dedicada de “saída rápida” foi removida do HTML (comentário: saídas via **Requisições**). O **Log de Saídas** lista registros do `history` filtrados para leitura operacional.

### 6.6 Requisições (`requisitions-view`)

Criação, listagem, numeração sequencial, detalhes, impressão/PDF (html2canvas + jsPDF), fluxo de status (pendente/aprovado etc. — ver strings no render). Integração com baixa de estoque ao concluir requisição (transações/batch no código).

### 6.7 Log de atividades (`activity-log-view`)

Visão mais ampla de auditoria (tipos de evento no histórico).

### 6.8 Baixa RM (`rm-view`)

Lista entradas do histórico com checkbox **`rmProcessed`**: marcação administrativa de que a baixa no sistema RM externo foi tratada. Atualiza documento no Firestore.

### 6.9 Cautela (`tool-loans-view`)

Registro de retirada de ferramentas com responsável, papel, observação; fila de várias linhas; baixa de estoque em **batch**; listas de abertas/devolvidas; export; assinatura em canvas (`setupSignaturePad`) onde previsto no HTML.

### 6.10 Gerar plaquinhas (`plaques-view`)

Geração de etiquetas/PDF com suporte a deep link (função `tryOpenPlaqueDeepLink` após carga).

### 6.11 Sugestão de compras (`compras-view`)

Análise de consumo/sazonalidade com período selecionável, categorias (pedido, giro, zero, baixo), export Excel com snapshot `lastComprasExport`.

### 6.12 Relatórios (`reports-view`)

Relatórios de saída, consumo, KPIs (botões ligados a `generate-exit-report-btn`, etc.).

### 6.13 Estoque Estrela (`estrela-view`)

Sub-abas: estoque próprio, entradas, saídas. Códigos sequenciais `EST-0001`, etc. CRUD em coleções `estrela_*`. Listeners só ativos quando essa view está ativa.

### 6.14 Configurações (modal)

Acesso por `#settings-btn` — restrito a admin (`settings`). Inclui nome do app, logo (URL), gestão de usuários (onde implementado), backup/restore, etc.

### 6.15 Feedback

Painel de sugestões grava em `.../suggestions` com `obraId` e metadados.

### 6.16 IA (Gemini / TTS)

Funções `callGeminiAPI` e `callTTSAPI`: a **API key está vazia** no código; sem chave, o app avisa via toast. O botão de IA na lista de produtos permanece oculto enquanto não houver integração configurada. **Não commite chaves** no repositório.

---

## 7. Backup e restauração

### 7.1 Formato do arquivo

Download: `backup-estoque-YYYY-MM-DD.json`.

Conteúdo:

```json
{
  "products": [ ... ],
  "history": [ ... ],
  "locations": [ ... ],
  "requisitions": [ ... ],
  "toolLoans": [ ... ],
  "settings": { ... }
}
```

Timestamps do Firestore são serializados como `{ "_fs_timestamp": true, "seconds": N, "nanoseconds": N }` para poder reconstruir `Timestamp` na restauração.

### 7.2 Restaurar

1. Apaga **todos** os documentos das coleções: products, history, locations, requisitions, tool_loans (da obra atual).
2. Recria documentos com os mesmos `id` quando presentes no backup.
3. Atualiza `app_settings/main` se `settings` existir no arquivo.
4. Recarrega a página após sucesso.

**Permissão:** exige `import` no papel (apenas **admin** na matriz atual).

**Cuidado:** não inclui por padrão as coleções `estrela_*` nem `users` — backup é do núcleo operacional da obra selecionada.

---

## 8. Convenções de dados (produto)

Campos usados em vários pontos (não necessariamente todos obrigatórios no Firestore):

- `name`, `code`, `codeRM`, `location`, `group`, `unit`
- `quantity`, `minQuantity`
- `photoUrl` ou similar (URL da foto)
- IDs de documento Firestore como `id` nos objetos em memória

Textos são frequentemente normalizados para maiúsculas em gravações (`toUpperText`, etc.).

---

## 9. Service Worker e cache

- Instalação faz prefetch de `index.html`, CSS, JS, manifest, ícones.
- Mensagem `SKIP_WAITING` para ativar nova versão do SW.
- Ao publicar nova versão, incremente `CACHE_NAME` em `sw.js` e, se necessário, query string `?v=` nos links de CSS/JS em `index.html` (já há `?v=2.0.4` nos links de estilo).

---

## 10. Dependências externas (CDN)

Listadas no `<head>` de `index.html`: Google Fonts (Inter, Material Symbols), html2canvas, jsPDF, Chart.js, xlsx-js-style, módulos Firebase 11.6.1.

---

## 11. Manutenção e extensão

### 11.1 Onde alterar o quê

| Objetivo | Arquivo principal |
|----------|-------------------|
| Novo campo de produto | `index.html` (form + modal edição) + `main.js` (submit, render, Firestore) |
| Nova aba / tela | `index.html` (bloco view) + sidebar/nav + `switchView` + listeners |
| Regras de permissão | `PERMISSIONS` e `updateUIBasedOnPermissions` |
| Caminhos Firestore | `initializeAppSession` (`obraBase`, refs) |
| Tema Tailwind / cores | `tailwind.config.js` + classes no HTML |

### 11.2 Testes

`package.json` não define suite de testes automatizados (`npm test` é placeholder). Testes manuais recomendados: login por papel, CRUD produto, entrada, requisição completa, cautela, backup/restore em ambiente de homologação.

### 11.3 Consistência de versão

Alinhar versão entre: rodapé da UI (`index.html`), `manifest.json`, `sw.js` (`CACHE_NAME`), query `?v=` dos assets e `package.json` ao lançar release.

---

## 12. Glossário rápido

- **RM:** código de material (referência externa / SAP-like).
- **Cautela:** empréstimo controlado de ferramenta com devolução.
- **Requisição:** pedido formal de material com trilha e PDF.
- **Obra:** tenant lógico (`uhe_estrela` vs outras pastas sob `obras/`).

---

## 13. Referências no código

Pontos de entrada úteis em `assets/js/main.js`:

- Config Firebase e `appId` — início do arquivo.
- `PERMISSIONS`, `hasPermission`, `hashPassword`, `normalizeUserId`.
- `initializeAppSession` — refs Firestore e UI pós-login.
- `startCoreListeners` / `startEstrelaListeners` / `syncRealtimeListeners`.
- `doLogin` / logout.
- `addHistoryEntry`, `switchView`.
- Backup/restore — handlers de `backupBtn` / `restoreFileInput`.
- Módulo Estrela — comentário `MÓDULO COMPLETO — Controle de Estoque UHE Estrela` em diante.

---

*Documento gerado com base na leitura do código-fonte do repositório. Ajuste este arquivo sempre que o comportamento do app mudar.*
