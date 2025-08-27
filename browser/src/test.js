/* eslint-disable no-global-assign */

// Mock browser storage to automatically enable the extension
browser = {
	storage: {
		sync: {
			get: async () => {
				return { enabled: true };
			},
			set: async () => true,
			onChanged: {
				addListener: async () => true,
			},
		},
	},
};
