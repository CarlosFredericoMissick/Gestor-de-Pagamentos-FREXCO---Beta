// ========================================
// CONFIGURAÇÃO — SharePoint / MSAL
// ========================================

const SP_SITE_ID = 'frexcocombr.sharepoint.com,985aeb05-f8f1-4931-a292-f020b17fdd52,899f116a-6492-4d7e-84b8-e8d1c208c628';

const SP_SCOPES = ['User.Read', 'Files.ReadWrite.AppFolder'];

// Pasta da aplicação no drive do site (prestadores.json, recebivel.json, notas/, pdfs/, contratos)
const GRAPH_APPROOT = `https://graph.microsoft.com/v1.0/sites/${SP_SITE_ID}/drive/special/approot:`;
const GRAPH_ITEMS   = `https://graph.microsoft.com/v1.0/sites/${SP_SITE_ID}/drive/items`;

const msalConfig = {
    auth: {
        clientId: '34f9f656-b81e-4698-abde-90293dd9d236',
        authority: 'https://login.microsoftonline.com/cc139011-64f1-4ea5-b3b0-2d5e3d8d503e',
        redirectUri: window.location.origin
    }
};
