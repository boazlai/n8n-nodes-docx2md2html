import type {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { NodeConnectionTypes, NodeOperationError } from 'n8n-workflow';
import mammoth from 'mammoth';
import { marked } from 'marked';
import HTMLtoDOCX from 'html-to-docx';

// mammoth.convertToMarkdown exists at runtime but is missing from the bundled types
type MammothFull = typeof mammoth & {
	convertToMarkdown: (
		input: { buffer: Buffer | ArrayBuffer } | { path: string } | { arrayBuffer: ArrayBuffer },
		options?: Parameters<typeof mammoth.convertToHtml>[1],
	) => ReturnType<typeof mammoth.convertToHtml>;
};

export class Docx2Md implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'HTML ↔ DOCX ↔ Markdown',
		name: 'docx2Md',
		icon: { light: 'file:docx2md.svg', dark: 'file:docx2md.dark.svg' },
		group: ['transform'],
		version: 1,
		description:
			'Convert DOCX files to Markdown or HTML, and convert Markdown or HTML back to DOCX',
		usableAsTool: true,
		defaults: {
			name: 'DOCX ↔ Markdown / HTML',
		},
		inputs: [NodeConnectionTypes.Main],
		outputs: [NodeConnectionTypes.Main],
		properties: [
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'DOCX → HTML',
						value: 'docxToHtml',
						description: 'Convert a DOCX binary file to HTML text',
						action: 'Convert DOCX to HTML',
					},
					{
						name: 'DOCX → Markdown',
						value: 'docxToMarkdown',
						description: 'Convert a DOCX binary file to Markdown text',
						action: 'Convert DOCX to markdown',
					},
					{
						name: 'HTML → DOCX',
						value: 'htmlToDocx',
						description: 'Convert HTML text to a DOCX binary file',
						action: 'Convert HTML to DOCX file',
					},
					{
						name: 'Markdown → DOCX',
						value: 'markdownToDocx',
						description: 'Convert Markdown text to a DOCX binary file',
						action: 'Convert Markdown to DOCX file',
					},
				],
				default: 'docxToMarkdown',
			},

			// ── DOCX → Markdown ──────────────────────────────────────────────
			{
				displayName: 'Input Binary Field',
				name: 'binaryPropertyName',
				type: 'string',
				default: 'data',
				required: true,
				displayOptions: { show: { operation: ['docxToMarkdown'] } },
				description: 'Name of the binary field that holds the DOCX file',
			},
			{
				displayName: 'Output JSON Field',
				name: 'outputMarkdownJsonField',
				type: 'string',
				default: 'markdown',
				required: true,
				displayOptions: { show: { operation: ['docxToMarkdown'] } },
				description: 'Name of the JSON property where the markdown text will be stored',
			},
			{
				displayName: 'Extract Images',
				name: 'extractImages',
				type: 'boolean',
				default: false,
				displayOptions: { show: { operation: ['docxToMarkdown'] } },
				description:
					'Whether to extract embedded images from the DOCX. When enabled the output JSON field becomes an object with "markdown" and "images" keys; image references in the markdown use ![](image_1) syntax.',
			},
			{
				displayName: 'Output Binary .md File',
				name: 'outputBinaryFile',
				type: 'boolean',
				default: false,
				displayOptions: { show: { operation: ['docxToMarkdown'] } },
				description:
					'Whether to also attach the markdown as a binary .md file (in addition to the JSON field)',
			},
			{
				displayName: 'Output Markdown Binary Field',
				name: 'outputMarkdownBinaryField',
				type: 'string',
				default: 'markdown',
				displayOptions: { show: { operation: ['docxToMarkdown'], outputBinaryFile: [true] } },
				description: 'Binary field name where the output .md file will be stored',
			},

			// ── Markdown → DOCX ──────────────────────────────────────────────
			{
				displayName: 'Markdown Source',
				name: 'markdownSource',
				type: 'options',
				options: [
					{
						name: 'Enter Markdown',
						value: 'editor',
						description: 'Type or paste Markdown directly into the node',
					},
					{
						name: 'JSON Text Field',
						value: 'textField',
						description: 'Read Markdown from a text property in the item JSON',
					},
					{
						name: 'Binary File',
						value: 'binaryFile',
						description: 'Read Markdown from an attached binary .md file',
					},
				],
				default: 'editor',
				displayOptions: { show: { operation: ['markdownToDocx'] } },
				description: 'Where to read the Markdown content from',
			},
			{
				displayName: 'Markdown',
				name: 'markdownContent',
				type: 'string',
				default: '',
				placeholder: '# Heading\n\nPaste your **Markdown** here...',
				required: true,
				displayOptions: {
					show: { operation: ['markdownToDocx'], markdownSource: ['editor'] },
				},
				description: 'The Markdown content to convert to DOCX',
			},
			{
				displayName: 'Markdown Field Name',
				name: 'markdownField',
				type: 'string',
				default: 'markdown',
				required: true,
				displayOptions: {
					show: { operation: ['markdownToDocx'], markdownSource: ['textField'] },
				},
				description: 'Name of the JSON property that contains the Markdown string',
			},
			{
				displayName: 'Input Binary Field',
				name: 'binaryPropertyNameMd',
				type: 'string',
				default: 'data',
				required: true,
				displayOptions: {
					show: { operation: ['markdownToDocx'], markdownSource: ['binaryFile'] },
				},
				description: 'Name of the binary field that holds the Markdown (.md) file',
			},
			{
				displayName: 'Output DOCX Binary Field',
				name: 'outputDocxBinaryField',
				type: 'string',
				default: 'data',
				required: true,
				displayOptions: { show: { operation: ['markdownToDocx'] } },
				description: 'Binary field name where the output .docx file will be stored',
			},
			{
				displayName: 'Output File Name',
				name: 'outputFileName',
				type: 'string',
				default: 'output.docx',
				displayOptions: { show: { operation: ['markdownToDocx'] } },
				description: 'File name for the generated DOCX file',
			},

			// ── DOCX → HTML ──────────────────────────────────────────────────
			{
				displayName: 'Input Binary Field',
				name: 'binaryPropertyNameDocxHtml',
				type: 'string',
				default: 'data',
				required: true,
				displayOptions: { show: { operation: ['docxToHtml'] } },
				description: 'Name of the binary field that holds the DOCX file',
			},
			{
				displayName: 'Output JSON Field',
				name: 'outputHtmlJsonField',
				type: 'string',
				default: 'html',
				required: true,
				displayOptions: { show: { operation: ['docxToHtml'] } },
				description: 'Name of the JSON property where the HTML text will be stored',
			},
			{
				displayName: 'Extract Images',
				name: 'extractImagesHtml',
				type: 'boolean',
				default: false,
				displayOptions: { show: { operation: ['docxToHtml'] } },
				description:
					'Whether to extract embedded images from the DOCX. When enabled the output JSON field becomes an object with "html" and "images" keys; image src attributes use image_1 keys instead of inline base64.',
			},
			{
				displayName: 'Output Binary .html File',
				name: 'outputHtmlBinaryFile',
				type: 'boolean',
				default: false,
				displayOptions: { show: { operation: ['docxToHtml'] } },
				description:
					'Whether to also attach the HTML as a binary .html file (in addition to the JSON field)',
			},
			{
				displayName: 'Output HTML Binary Field',
				name: 'outputHtmlBinaryField',
				type: 'string',
				default: 'html',
				displayOptions: { show: { operation: ['docxToHtml'], outputHtmlBinaryFile: [true] } },
				description: 'Binary field name where the output .html file will be stored',
			},

			// ── HTML → DOCX ──────────────────────────────────────────────────
			{
				displayName: 'HTML Source',
				name: 'htmlSource',
				type: 'options',
				options: [
					{
						name: 'Enter HTML',
						value: 'editor',
						description: 'Type or paste HTML directly into the node',
					},
					{
						name: 'JSON Text Field',
						value: 'textField',
						description: 'Read HTML from a text property in the item JSON',
					},
					{
						name: 'Binary File',
						value: 'binaryFile',
						description: 'Read HTML from an attached binary .html file',
					},
				],
				default: 'editor',
				displayOptions: { show: { operation: ['htmlToDocx'] } },
				description: 'Where to read the HTML content from',
			},
			{
				displayName: 'HTML',
				name: 'htmlContent',
				type: 'string',
				default: '',
				placeholder: '<h1>Heading</h1>\n\n<p>Paste your <strong>HTML</strong> here...</p>',
				required: true,
				displayOptions: {
					show: { operation: ['htmlToDocx'], htmlSource: ['editor'] },
				},
				description: 'The HTML content to convert to DOCX',
			},
			{
				displayName: 'HTML Field Name',
				name: 'htmlField',
				type: 'string',
				default: 'html',
				required: true,
				displayOptions: {
					show: { operation: ['htmlToDocx'], htmlSource: ['textField'] },
				},
				description: 'Name of the JSON property that contains the HTML string',
			},
			{
				displayName: 'Input Binary Field',
				name: 'binaryPropertyNameHtml',
				type: 'string',
				default: 'data',
				required: true,
				displayOptions: {
					show: { operation: ['htmlToDocx'], htmlSource: ['binaryFile'] },
				},
				description: 'Name of the binary field that holds the HTML file',
			},
			{
				displayName: 'Output DOCX Binary Field',
				name: 'outputDocxBinaryFieldHtml',
				type: 'string',
				default: 'data',
				required: true,
				displayOptions: { show: { operation: ['htmlToDocx'] } },
				description: 'Binary field name where the output .docx file will be stored',
			},
			{
				displayName: 'Output File Name',
				name: 'outputFileNameHtml',
				type: 'string',
				default: 'output.docx',
				displayOptions: { show: { operation: ['htmlToDocx'] } },
				description: 'File name for the generated DOCX file',
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];

		for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
			try {
				const operation = this.getNodeParameter('operation', itemIndex) as string;

				if (operation === 'docxToMarkdown') {
					const inputField = this.getNodeParameter(
						'binaryPropertyName',
						itemIndex,
						'data',
					) as string;
					const outputBinaryFile = this.getNodeParameter(
						'outputBinaryFile',
						itemIndex,
						false,
					) as boolean;
					const outputJsonField = this.getNodeParameter(
						'outputMarkdownJsonField',
						itemIndex,
						'markdown',
					) as string;

					const binaryMeta = this.helpers.assertBinaryData(itemIndex, inputField);
					const docxBuffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputField);

					const extractImages = this.getNodeParameter('extractImages', itemIndex, false) as boolean;

					let markdownText: string;
					const collectedImages: Array<Record<string, string>> = [];

					if (extractImages) {
						let imageCounter = 0;
						const { value } = await (mammoth as MammothFull).convertToMarkdown(
							{ buffer: docxBuffer },
							{
								convertImage: mammoth.images.imgElement(async (image) => {
									imageCounter++;
									const key = `image_${imageCounter}`;
									const b64 = await image.readAsBase64String();
									collectedImages.push({ [key]: b64 });
									return { src: key };
								}),
							},
						);
						markdownText = value;
					} else {
						const { value } = await (mammoth as MammothFull).convertToMarkdown({
							buffer: docxBuffer,
						});
						markdownText = value;
					}

					const jsonValue = extractImages
						? { markdown: markdownText, images: collectedImages }
						: markdownText;

					const outputItem: INodeExecutionData = {
						json: { ...items[itemIndex].json, [outputJsonField]: jsonValue },
						pairedItem: itemIndex,
					};

					if (outputBinaryFile) {
						const outputField = this.getNodeParameter(
							'outputMarkdownBinaryField',
							itemIndex,
							'markdown',
						) as string;
						const baseName = (binaryMeta.fileName ?? 'document').replace(/\.docx$/i, '');
						const mdBinary = await this.helpers.prepareBinaryData(
							Buffer.from(markdownText, 'utf-8'),
							`${baseName}.md`,
							'text/markdown',
						);
						outputItem.binary = { ...(items[itemIndex].binary ?? {}), [outputField]: mdBinary };
					}

					returnData.push(outputItem);
				} else if (operation === 'docxToHtml') {
					const inputField = this.getNodeParameter(
						'binaryPropertyNameDocxHtml',
						itemIndex,
						'data',
					) as string;
					const outputBinaryFile = this.getNodeParameter(
						'outputHtmlBinaryFile',
						itemIndex,
						false,
					) as boolean;
					const outputJsonField = this.getNodeParameter(
						'outputHtmlJsonField',
						itemIndex,
						'html',
					) as string;

					const binaryMeta = this.helpers.assertBinaryData(itemIndex, inputField);
					const docxBuffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputField);

					const extractImages = this.getNodeParameter(
						'extractImagesHtml',
						itemIndex,
						false,
					) as boolean;

					let htmlText: string;
					const collectedImages: Array<Record<string, string>> = [];

					if (extractImages) {
						let imageCounter = 0;
						const { value } = await mammoth.convertToHtml(
							{ buffer: docxBuffer },
							{
								convertImage: mammoth.images.imgElement(async (image) => {
									imageCounter++;
									const key = `image_${imageCounter}`;
									const b64 = await image.readAsBase64String();
									collectedImages.push({ [key]: b64 });
									return { src: key };
								}),
							},
						);
						htmlText = value;
					} else {
						const { value } = await mammoth.convertToHtml({ buffer: docxBuffer });
						htmlText = value;
					}

					const jsonValue = extractImages ? { html: htmlText, images: collectedImages } : htmlText;

					const outputItem: INodeExecutionData = {
						json: { ...items[itemIndex].json, [outputJsonField]: jsonValue },
						pairedItem: itemIndex,
					};

					if (outputBinaryFile) {
						const outputField = this.getNodeParameter(
							'outputHtmlBinaryField',
							itemIndex,
							'html',
						) as string;
						const baseName = (binaryMeta.fileName ?? 'document').replace(/\.docx$/i, '');
						const htmlBinary = await this.helpers.prepareBinaryData(
							Buffer.from(htmlText, 'utf-8'),
							`${baseName}.html`,
							'text/html',
						);
						outputItem.binary = { ...(items[itemIndex].binary ?? {}), [outputField]: htmlBinary };
					}

					returnData.push(outputItem);
				} else if (operation === 'markdownToDocx') {
					const markdownSource = this.getNodeParameter(
						'markdownSource',
						itemIndex,
						'editor',
					) as string;
					const outputField = this.getNodeParameter(
						'outputDocxBinaryField',
						itemIndex,
						'data',
					) as string;
					const outputFileName = this.getNodeParameter(
						'outputFileName',
						itemIndex,
						'output.docx',
					) as string;

					let markdownText: string;

					if (markdownSource === 'editor') {
						markdownText = this.getNodeParameter('markdownContent', itemIndex, '') as string;
					} else if (markdownSource === 'textField') {
						const fieldName = this.getNodeParameter(
							'markdownField',
							itemIndex,
							'markdown',
						) as string;
						const fieldValue = items[itemIndex].json[fieldName];
						if (typeof fieldValue !== 'string') {
							throw new NodeOperationError(
								this.getNode(),
								`Field "${fieldName}" does not contain a string value`,
								{ itemIndex },
							);
						}
						markdownText = fieldValue;
					} else {
						const inputField = this.getNodeParameter(
							'binaryPropertyNameMd',
							itemIndex,
							'data',
						) as string;
						const mdBuffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputField);
						markdownText = mdBuffer.toString('utf-8');
					}

					const htmlFromMd = marked.parse(markdownText) as string;
					const docxResultMd = await HTMLtoDOCX(htmlFromMd, null, {
						table: { row: { cantSplit: true } },
						footer: false,
						pageNumber: false,
					});

					const docxBufferMd = Buffer.isBuffer(docxResultMd)
						? docxResultMd
						: Buffer.from(docxResultMd as ArrayBuffer);

					const docxBinaryMd = await this.helpers.prepareBinaryData(
						docxBufferMd,
						outputFileName,
						'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
					);

					returnData.push({
						json: { ...items[itemIndex].json },
						binary: { ...(items[itemIndex].binary ?? {}), [outputField]: docxBinaryMd },
						pairedItem: itemIndex,
					});
				} else {
					// htmlToDocx
					const htmlSource = this.getNodeParameter('htmlSource', itemIndex, 'editor') as string;
					const outputField = this.getNodeParameter(
						'outputDocxBinaryFieldHtml',
						itemIndex,
						'data',
					) as string;
					const outputFileName = this.getNodeParameter(
						'outputFileNameHtml',
						itemIndex,
						'output.docx',
					) as string;

					let htmlText: string;

					if (htmlSource === 'editor') {
						htmlText = this.getNodeParameter('htmlContent', itemIndex, '') as string;
					} else if (htmlSource === 'textField') {
						const fieldName = this.getNodeParameter('htmlField', itemIndex, 'html') as string;
						const fieldValue = items[itemIndex].json[fieldName];
						if (typeof fieldValue !== 'string') {
							throw new NodeOperationError(
								this.getNode(),
								`Field "${fieldName}" does not contain a string value`,
								{ itemIndex },
							);
						}
						htmlText = fieldValue;
					} else {
						const inputField = this.getNodeParameter(
							'binaryPropertyNameHtml',
							itemIndex,
							'data',
						) as string;
						const htmlBuffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputField);
						htmlText = htmlBuffer.toString('utf-8');
					}

					const docxResultHtml = await HTMLtoDOCX(htmlText, null, {
						table: { row: { cantSplit: true } },
						footer: false,
						pageNumber: false,
					});

					const docxBufferHtml = Buffer.isBuffer(docxResultHtml)
						? docxResultHtml
						: Buffer.from(docxResultHtml as ArrayBuffer);

					const docxBinaryHtml = await this.helpers.prepareBinaryData(
						docxBufferHtml,
						outputFileName,
						'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
					);

					returnData.push({
						json: { ...items[itemIndex].json },
						binary: { ...(items[itemIndex].binary ?? {}), [outputField]: docxBinaryHtml },
						pairedItem: itemIndex,
					});
				}
			} catch (error) {
				if (this.continueOnFail()) {
					returnData.push({
						json: this.getInputData(itemIndex)[0].json,
						error,
						pairedItem: itemIndex,
					});
				} else {
					if (error.context) {
						error.context.itemIndex = itemIndex;
						throw error;
					}
					throw new NodeOperationError(this.getNode(), error as Error, { itemIndex });
				}
			}
		}

		return [returnData];
	}
}
