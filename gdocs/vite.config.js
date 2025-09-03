import { defineConfig } from 'vite';
import { resolve } from 'path';
import { viteSingleFile } from 'vite-plugin-singlefile';
import { viteStaticCopy } from 'vite-plugin-static-copy';

const targets = [
	{ src: 'server/*', dest: 'server' },
	{ src: '*.json', dest: '.' },
];

export default defineConfig({
	root: 'src',
	plugins: [viteSingleFile(), viteStaticCopy({ targets })],
	build: {
		outDir: resolve(__dirname, 'dist'),
		rollupOptions: {
			input: resolve(__dirname, 'src/ui/sidebar.html'),
		},
		minify: false,
	},
	server: {
		port: 3000,
		open: true,
	},
});
