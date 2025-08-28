import { defineConfig } from 'vite';

export default defineConfig({
	test: {
		globals: true,
		environment: 'jsdom',
		setupFiles: './setupTests.js',
		coverage: {
			reporter: ['text', 'html'],
			exclude: ['node_modules/, setupTests.js'],
		},
	},
});
