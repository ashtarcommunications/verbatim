import nsda from '@speechanddebate/eslint-config-nsda';

export default [
	...nsda,
	{
		languageOptions: {
			globals: {
				browser: 'readonly',
				hotkeys: 'readonly',
				splitter: 'readonly',
			},
		},
	},
	{
		rules: {},
	},
];
