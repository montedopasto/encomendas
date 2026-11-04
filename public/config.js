/*
 * Configuração da instalação da app Encomendas. Preencher pelo TI (ver docs/INSTALAR.md).
 * Nada aqui é secreto: o clientId de uma app de página (SPA) é público por natureza.
 */
window.ENCOMENDAS_CONFIG = {
  // Registo da app no Microsoft Entra ID (Azure AD) — "Aplicação de página única (SPA)".
  tenantId: 'COLOCAR-ID-DO-TENANT',          // ID do diretório (tenant) da Monte do Pasto
  clientId: 'COLOCAR-ID-DA-APLICACAO',       // ID da aplicação (cliente)
  redirectUri: '',                           // vazio = endereço da própria página (tem de estar registado no Entra)

  // Só entram contas deste domínio.
  domain: 'montedopasto.pt',

  // Pasta onde ficam os dados da app (encomendas, clientes, stock, definições), em ficheiros JSON.
  // Sugestão: SharePoint › Documentos Partilhados › Pedro Sá › IA › SalesVision › dados
  // A biblioteca (drive) é a de "Documentos Partilhados" do site Monte do Pasto.
  dataDriveId: 'b!tDspHS-8_06fSneRoEVGiIzQ_wwWK5dJhqGpmrL0k_yVQ7fm29iyQ7jbVXGCuMch',
  dataFolderId: '01NEJB4KDNHFIF25CJLVEKJDMOHLMRTL6U', // pasta "dados" criada a 02/11/2026

  // Quem pode alterar as Definições e importar cópias (emails em minúsculas). Lista vazia = todos.
  admins: ['pedro.sa@montedopasto.pt'],

  // De quantos em quantos milissegundos a app vai buscar alterações feitas por outras pessoas.
  pollMs: 15000
};
