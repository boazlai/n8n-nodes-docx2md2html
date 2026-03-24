// Augment mammoth to add convertToMarkdown, which exists at runtime but is not
// included in the bundled type declarations.
import 'mammoth';

declare module 'mammoth' {
	interface Mammoth {
		convertToMarkdown: (input: Input, options?: Options) => Promise<Result>;
	}
}
