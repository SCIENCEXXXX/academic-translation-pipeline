
module.exports = {
	stylesheet: './pdf-style.css',
	body_class: 'markdown-body',
	pdf_options: {
		format: 'A4',
		margin: '20mm',
		printBackground: true,
		displayHeaderFooter: true,
		footerTemplate: `
			<div style="font-size: 8px; margin: 0 auto; color: #999;">
				<span class="pageNumber"></span> / <span class="totalPages"></span>
			</div>`,
		headerTemplate: '<div style="font-size: 8px; margin: 0 auto; color: #999;">生物力学训练与测试</div>',
	},
	launch_options: {
		args: ['--no-sandbox', '--disable-setuid-sandbox'],
	},
	// This helps with generating TOC/Outline in some Puppeteer versions
	devtools: false,
	// Inject KaTeX for math rendering and wait for it
	script: {
		content: `
		(async function() {
			// Add KaTeX CSS
			const link = document.createElement('link');
			link.rel = 'stylesheet';
			link.href = 'https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/katex.min.css';
			document.head.appendChild(link);

			// Add KaTeX JS
			const script1 = document.createElement('script');
			script1.src = 'https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/katex.min.js';
			document.head.appendChild(script1);

			await new Promise(resolve => script1.onload = resolve);

			// Add KaTeX Auto-render JS
			const script2 = document.createElement('script');
			script2.src = 'https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/contrib/auto-render.min.js';
			document.head.appendChild(script2);

			await new Promise(resolve => script2.onload = resolve);

			// Render math
			renderMathInElement(document.body, {
				delimiters: [
					{left: '$$', right: '$$', display: true},
					{left: '$', right: '$', display: false}
				],
				throwOnError: false
			});

			// Mark as finished for md-to-pdf to capture
			window.status = 'ready';
		})();
		`
	},
};
