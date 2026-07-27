import OAuthProvider from '@cloudflare/workers-oauth-provider';
import { McpAgent } from 'agents/mcp';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@larksuiteoapi/node-sdk';
import { env } from 'cloudflare:workers';

import { FeishuHandler } from './feishu-handler';
import { Props, refreshUpstreamAuthToken } from './utils';
import { oapiHttpInstance } from './utils/http-instance';
import { RecallTool } from './mcp-tool/document-tool/recall';
import { blockTreeTool, docxBlockBatchDelete, docxBlockPatch, docxInsertImage, docxInsertFile, docxMarkdownImport, docxV1BlockTypeSchemaGet, docxV1DocumentBlockChildrenCreateSimple } from './tools/document';
import { mediaUploadTool } from './tools/drive';
import { z } from 'zod';
import { driveCommentBatch, driveCommentList,driveCommentPatch,driveCommentCreate,driveCommentGet } from './tools/drive/comment';
import { driveReplyList, driveReplyUpdate, driveReplyDelete } from './tools/drive/reply';
import { wikiNodeInfoGet } from './tools/wiki/space';
import { sheetRangeRead, sheetInfoGet,sheetPatch, sheetRangeWrite } from './tools/sheet';
import { suiteSearch } from './tools/suite';
import { docxMarkdownInsert } from './tools/document';

import {
  registerTools,
  // authen
  getUserInfo,
  // docx
  createDocument,
  getDocument,
  getDocumentRawContent,
  convertContentToBlocks,
  // docx blocks
  listDocumentBlocks,
  createBlocks,
  deleteBlock,
  batchDeleteBlocks,
  buildTextBlock,
  buildHeading1Block,
  buildHeading2Block,
  buildHeading3Block,
  buildHeading4Block,
  buildHeading5Block,
  buildHeading6Block,
  buildHeading7Block,
  buildHeading8Block,
  buildHeading9Block,
  buildBulletBlock,
  buildOrderedBlock,
  buildQuoteBlock,
  buildEquationBlock,
  buildTodoBlock,
  buildCodeBlock,
  buildDividerBlock,
  buildCalloutBlock,
  searchFeishuCalloutEmoji,
  createFileBlock,
  createImageBlock,
  buildIframeBlock,
  buildChatCardBlock,
  buildGridBlock,
  buildMermaidBlock,
  buildGlossaryBlock,
  buildTimelineBlock,
  buildCatalogNavigationBlock,
  buildInformationCollectionBlock,
  buildCountdownBlock,
  // drive
  listFileComments,
  // sheets
  addSheet,
  copySheet,
  createSpreadsheet,
  deleteSheet,
  getSheet,
  getSpreadsheet,
  querySheets,
  updateSheetMetadata,
  updateSheetProtection,
  updateSheetViewSettings,
  updateSpreadsheet,
} from 'feishu-tools';

import { GenTools } from './mcp-tool/tools/zh/gen-tools';

const client = new Client({
  appId: env.FEISHU_APP_ID,
  appSecret: env.FEISHU_APP_SECRET,
  httpInstance: oapiHttpInstance,
});
export class MyMCP extends McpAgent<Props, Env> {
  server = new McpServer({
    name: 'Feishu OAuth Proxy Demo',
    version: '1.0.0',
  });
  // 统一回调，打包params和client、userAccessToken，传给customHandler
  async handler(params: any, customHandler: any) {
    return await customHandler(params, client, this.props.accessToken);
  }

  async init() {
    // Use the upstream access token to facilitate tools
    const context = {
      client: client,
      getUserAccessToken: () => this.props.accessToken as string,
    }

    // 批量注册所有 feishu-tools 工具
    const allTools = [
      // authen
      getUserInfo,
      // docx
      createDocument,
      getDocument,
      getDocumentRawContent,
      convertContentToBlocks,
      // docx blocks
      listDocumentBlocks,
      createBlocks,
      deleteBlock,
      batchDeleteBlocks,
      buildTextBlock,
      buildHeading1Block,
      buildHeading2Block,
      buildHeading3Block,
      buildHeading4Block,
      buildHeading5Block,
      buildHeading6Block,
      buildHeading7Block,
      buildHeading8Block,
      buildHeading9Block,
      buildBulletBlock,
      buildOrderedBlock,
      buildQuoteBlock,
      buildEquationBlock,
      buildTodoBlock,
      buildCodeBlock,
      buildDividerBlock,
      buildCalloutBlock,
      searchFeishuCalloutEmoji,
      createFileBlock,
      createImageBlock,
      buildIframeBlock,
      buildChatCardBlock,
      buildGridBlock,
      buildMermaidBlock,
      buildGlossaryBlock,
      buildTimelineBlock,
      buildCatalogNavigationBlock,
      buildInformationCollectionBlock,
      buildCountdownBlock,
      // drive
      listFileComments,
      // sheets
      addSheet,
      copySheet,
      createSpreadsheet,
      deleteSheet,
      getSheet,
      getSpreadsheet,
      querySheets,
      updateSheetMetadata,
      updateSheetProtection,
      updateSheetViewSettings,
      updateSpreadsheet,
    ];

    registerTools(this.server, allTools, context);

    // Generic Drive tools intentionally use the current Feishu user's OAuth token.
    // They never accept a server-side filesystem path: files are supplied as base64
    // so an MCP client cannot make the Worker read arbitrary local files.
    const requestDrive = async (path: string, init: RequestInit = {}) => {
      const response = await fetch(`https://open.feishu.cn/open-apis${path}`, {
        ...init,
        headers: {
          Authorization: `Bearer ${this.props.accessToken}`,
          ...(init.headers || {}),
        },
      });
      const payload = await response.json() as { code?: number; msg?: string; data?: unknown };
      if (!response.ok || payload.code !== 0) {
        throw new Error(`Feishu Drive request failed: ${payload.msg || response.statusText} (code ${payload.code ?? response.status})`);
      }
      return payload.data;
    };

    const asResult = (data: unknown) => ({
      content: [{ type: 'text' as const, text: JSON.stringify(data, null, 2) }],
    });

    this.server.tool(
      'drive_get_root_folder',
      '获取当前已授权飞书用户的云空间根目录 token。后续列目录或创建文件夹时使用。',
      {},
      async () => asResult(await requestDrive('/drive/explorer/v2/root_folder/meta')),
    );

    this.server.tool(
      'drive_list_files',
      '列出飞书云盘指定文件夹中的文件和子文件夹。folder_token 可传 root；默认返回最多 100 项。',
      {
        folder_token: z.string().default('root').describe('要列出的文件夹 token；根目录用 root'),
        page_size: z.number().int().min(1).max(100).default(100).describe('单页返回数量'),
        page_token: z.string().optional().describe('上一页结果返回的分页 token'),
      },
      async ({ folder_token, page_size, page_token }) => {
        const query = new URLSearchParams({ folder_token, page_size: String(page_size) });
        if (page_token) query.set('page_token', page_token);
        return asResult(await requestDrive(`/drive/v1/files?${query.toString()}`));
      },
    );

    this.server.tool(
      'drive_create_folder',
      '在飞书云盘指定父文件夹中新建一个文件夹。不会覆盖同名文件夹。',
      {
        name: z.string().min(1).max(250).describe('新文件夹名称'),
        parent_folder_token: z.string().default('root').describe('父文件夹 token；根目录用 root'),
      },
      async ({ name, parent_folder_token }) => asResult(await requestDrive('/drive/v1/files/create_folder', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=utf-8' },
        body: JSON.stringify({ name, folder_token: parent_folder_token }),
      })),
    );

    this.server.tool(
      'drive_upload_file',
      '上传一个不超过 20MB 的普通文件到飞书云盘文件夹。内容必须是 base64 编码；同名文件不会自动覆盖。',
      {
        name: z.string().min(1).max(250).describe('包含扩展名的文件名'),
        folder_token: z.string().describe('目标文件夹 token'),
        content_base64: z.string().min(1).describe('文件的 Base64 内容；可带 data: 前缀'),
      },
      async ({ name, folder_token, content_base64 }) => {
        const encoded = content_base64.replace(/^data:[^;]+;base64,/, '');
        const binary = atob(encoded);
        const bytes = Uint8Array.from(binary, char => char.charCodeAt(0));
        if (bytes.byteLength > 20 * 1024 * 1024) {
          throw new Error('文件超过飞书全量上传 20MB 限制，请改用分片上传。');
        }
        const form = new FormData();
        form.append('file_name', name);
        form.append('parent_type', 'explorer');
        form.append('parent_node', folder_token);
        form.append('size', String(bytes.byteLength));
        form.append('file', new File([bytes], name));
        return asResult(await requestDrive('/drive/v1/files/upload_all', {
          method: 'POST',
          body: form,
        }));
      },
    );

    // MCP clients can have a much smaller message limit than Feishu's 20MB
    // upload limit. These three tools use Durable Object storage as a short-lived
    // staging area so clients can stream a local file in safe base64 chunks.
    const uploadPrefix = 'drive-upload:';
    type PendingUpload = { name: string; folderToken: string };

    this.server.tool(
      'drive_upload_begin',
      '开始一次分块上传，适合内容无法一次放入 MCP 消息的大文件。返回 transfer_id，随后按顺序追加分块。',
      {
        name: z.string().min(1).max(250).describe('包含扩展名的文件名'),
        folder_token: z.string().describe('目标文件夹 token'),
      },
      async ({ name, folder_token }) => {
        const transferId = crypto.randomUUID();
        await this.ctx.storage.put<PendingUpload>(`${uploadPrefix}${transferId}:meta`, {
          name,
          folderToken: folder_token,
        });
        return asResult({ transfer_id: transferId, max_chunk_bytes: 24576 });
      },
    );

    this.server.tool(
      'drive_upload_append',
      '向已开始的分块上传追加一个 Base64 分块。每块解码后不得超过 24KB，chunk_index 必须从 0 开始连续递增。',
      {
        transfer_id: z.string().uuid().describe('drive_upload_begin 返回的 transfer_id'),
        chunk_index: z.number().int().min(0).describe('从 0 开始的连续分块序号'),
        content_base64: z.string().min(1).max(40000).describe('不带 data: 前缀的 Base64 分块内容'),
      },
      async ({ transfer_id, chunk_index, content_base64 }) => {
        const clean = content_base64.replace(/^data:[^;]+;base64,/, '');
        const decodedSize = atob(clean).length;
        if (decodedSize > 24576) throw new Error('分块超过 24KB 限制。');
        const meta = await this.ctx.storage.get<PendingUpload>(`${uploadPrefix}${transfer_id}:meta`);
        if (!meta) throw new Error('上传会话不存在或已完成，请重新开始上传。');
        const key = `${uploadPrefix}${transfer_id}:part:${String(chunk_index).padStart(6, '0')}`;
        await this.ctx.storage.put(key, clean);
        return asResult({ transfer_id, chunk_index, received_bytes: decodedSize });
      },
    );

    this.server.tool(
      'drive_upload_complete',
      '完成分块上传：在 Worker 内合并已上传的分块并上传到飞书云盘。成功后会自动清理临时分块。',
      { transfer_id: z.string().uuid().describe('drive_upload_begin 返回的 transfer_id') },
      async ({ transfer_id }) => {
        const metaKey = `${uploadPrefix}${transfer_id}:meta`;
        const meta = await this.ctx.storage.get<PendingUpload>(metaKey);
        if (!meta) throw new Error('上传会话不存在或已完成，请重新开始上传。');
        const chunks = await this.ctx.storage.list<string>({ prefix: `${uploadPrefix}${transfer_id}:part:` });
        if (chunks.size === 0) throw new Error('未收到任何分块，无法完成上传。');
        const buffers = [...chunks.entries()]
          .sort(([left], [right]) => left.localeCompare(right))
          .map(([, encoded]) => Uint8Array.from(atob(encoded), char => char.charCodeAt(0)));
        const size = buffers.reduce((total, bytes) => total + bytes.byteLength, 0);
        if (size > 20 * 1024 * 1024) throw new Error('文件超过飞书全量上传 20MB 限制。');
        const content = new Uint8Array(size);
        let offset = 0;
        for (const buffer of buffers) {
          content.set(buffer, offset);
          offset += buffer.byteLength;
        }
        const form = new FormData();
        form.append('file_name', meta.name);
        form.append('parent_type', 'explorer');
        form.append('parent_node', meta.folderToken);
        form.append('size', String(size));
        form.append('file', new File([content], meta.name));
        const result = await requestDrive('/drive/v1/files/upload_all', { method: 'POST', body: form });
        await this.ctx.storage.delete([metaKey, ...chunks.keys()]);
        return asResult({ ...(result as object), file_name: meta.name, file_size: size });
      },
    );

    this.server.tool(
      'drive_move_file',
      '将飞书云盘中的文件或文件夹移动到目标文件夹。该操作会改变云端位置。',
      {
        file_token: z.string().describe('要移动的文件或文件夹 token'),
        type: z.enum(['file', 'folder', 'doc', 'docx', 'sheet', 'bitable', 'mindnote']).describe('资源类型'),
        destination_folder_token: z.string().describe('目标文件夹 token'),
      },
      async ({ file_token, type, destination_folder_token }) => asResult(await requestDrive(`/drive/v1/files/${encodeURIComponent(file_token)}/move`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=utf-8' },
        body: JSON.stringify({ type, folder_token: destination_folder_token }),
      })),
    );
  }
}

export default new OAuthProvider({
  apiHandlers: {
    '/mcp': MyMCP.serve('/mcp'),
    '/sse': MyMCP.serveSSE('/sse'),
  },
  defaultHandler: FeishuHandler,
  authorizeEndpoint: '/authorize',
  tokenEndpoint: '/token',
  clientRegistrationEndpoint: '/register',
  tokenExchangeCallback: async (options) => {
    if (options.grantType === 'authorization_code') {
      return {
        accessTokenProps: options.props,
        accessTokenTTL: options.props.expiresIn,
      };
    }
    if (options.grantType === 'refresh_token') {
      const [accessToken, refreshToken, expiresIn, errResponse] = await refreshUpstreamAuthToken({
        refreshToken: options.props.refreshToken,
        upstream_url: 'https://open.feishu.cn/open-apis/authen/v2/oauth/token',
        client_id: env.FEISHU_APP_ID,
        client_secret: env.FEISHU_APP_SECRET,
      });
      // console.log('refreshUpstreamAuthToken', refreshToken)
      if (!errResponse){
        return {
          newProps: {
            ...options.props,
            accessToken: accessToken,
            refreshToken: refreshToken,
          },
          accessTokenTTL: expiresIn,
        };
      }
    }
  },
});
