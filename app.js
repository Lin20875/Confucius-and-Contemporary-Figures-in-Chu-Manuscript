// 构建目录（支持 H1-H3）
function buildToc() {
	const article = document.getElementById('article');
	const headings = article.querySelectorAll('h1, h2, h3');
	const toc = document.getElementById('toc');
	const scopeSelect = document.getElementById('scope-select');

	toc.innerHTML = '';

	const ul = document.createElement('ul');
	const tocItems = [];
	let currentH1 = null;
	let currentH2 = null;

	headings.forEach((h) => {
		const level = Number(h.tagName.substring(1));
		const id = h.id || h.textContent.trim().replace(/\s+/g, '-').toLowerCase();
		if (!h.id) h.id = id;

		const li = document.createElement('li');
		li.className = `lvl-${level - 0}`; // lvl-1, lvl-2, lvl-3

		const a = document.createElement('a');
		a.href = `#${id}`;
		a.textContent = h.textContent;
		a.dataset.targetId = id;

		li.appendChild(a);
		ul.appendChild(li);

		// 维护层级树，用于范围计算
		const node = { id, level, el: h, parent: null, children: [] };
		if (level === 1) {
			currentH1 = node;
			currentH2 = null;
			tocItems.push(node);
		} else if (level === 2) {
			if (currentH1) {
				node.parent = currentH1;
				currentH1.children.push(node);
			} else {
				tocItems.push(node);
			}
			currentH2 = node;
		} else if (level === 3) {
			if (currentH2) {
				node.parent = currentH2;
				currentH2.children.push(node);
			} else if (currentH1) {
				node.parent = currentH1;
				currentH1.children.push(node);
			} else {
				tocItems.push(node);
			}
		}
	});

	toc.appendChild(ul);

	// 填充范围下拉（全部 + 所有标题）
	scopeSelect.innerHTML = '<option value="__ALL__">全部章节</option>';
	headings.forEach((h) => {
		const id = h.id;
		const level = Number(h.tagName.substring(1));
		const indent = level === 1 ? '' : (level === 2 ? '— ' : '—— ');
		const opt = document.createElement('option');
		opt.value = id;
		opt.textContent = `${indent}${h.textContent}`;
		scopeSelect.appendChild(opt);
	});
}

// 目录点击平滑滚动
function enableTocClick() {
	const toc = document.getElementById('toc');
	toc.addEventListener('click', (e) => {
		const a = e.target.closest('a');
		if (!a) return;
		e.preventDefault();
		const id = a.dataset.targetId;
		const target = document.getElementById(id);
		if (target) {
			target.scrollIntoView({ behavior: 'smooth', block: 'start' });
			history.replaceState(null, '', `#${id}`);
		}
	});
}

// 滚动高亮当前目录项
function observeHeadings() {
	// 如已有观察器，先断开，避免重复
	if (window.__headingObserver) {
		try { window.__headingObserver.disconnect(); } catch (_) {}
	}
	const toc = document.getElementById('toc');
	const links = Array.from(toc.querySelectorAll('a'));
	const linkMap = new Map(links.map((l) => [l.dataset.targetId, l.parentElement]));

	const observer = new IntersectionObserver((entries) => {
		entries.forEach((entry) => {
			const id = entry.target.id;
			const li = linkMap.get(id);
			if (!li) return;
			if (entry.isIntersecting) {
				links.forEach((l) => l.parentElement.classList.remove('active'));
				li.classList.add('active');
			}
		});
	}, {
		root: document.querySelector('.content'),
		threshold: 0.3
	});

	document.querySelectorAll('#article h1, #article h2, #article h3').forEach((h) => observer.observe(h));
	window.__headingObserver = observer;
}

// 清除高亮
function clearHighlights(root) {
	const marks = root.querySelectorAll('mark.__match');
	marks.forEach((mark) => {
		const parent = mark.parentNode;
		parent.replaceChild(document.createTextNode(mark.textContent), mark);
		parent.normalize();
	});
}

// 在元素内高亮文本
function highlightText(root, query) {
	if (!query) return 0;
	let count = 0;
	const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
		acceptNode(node) {
			if (!node.nodeValue.trim()) return NodeFilter.FILTER_REJECT;
			// 跳过已在 mark 中的文本
			if (node.parentElement && node.parentElement.tagName === 'MARK') return NodeFilter.FILTER_REJECT;
			return NodeFilter.FILTER_ACCEPT;
		}
	});

	const regex = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
	const textNodes = [];
	while (walker.nextNode()) textNodes.push(walker.currentNode);

	textNodes.forEach((textNode) => {
		const text = textNode.nodeValue;
		if (!regex.test(text)) return;
		const frag = document.createDocumentFragment();
		let lastIndex = 0;
		text.replace(regex, (match, idx) => {
			if (idx > lastIndex) {
				frag.appendChild(document.createTextNode(text.slice(lastIndex, idx)));
			}
			const mark = document.createElement('mark');
			mark.className = '__match';
			mark.textContent = match;
			frag.appendChild(mark);
			count += 1;
			lastIndex = idx + match.length;
			return match;
		});
		if (lastIndex < text.length) {
			frag.appendChild(document.createTextNode(text.slice(lastIndex)));
		}
		textNode.parentNode.replaceChild(frag, textNode);
	});
	return count;
}

// 计算某标题的“同级范围”（到下一个同级标题前）
function getHeadingScopeElement(heading) {
	const level = Number(heading.tagName.substring(1));
	const range = document.createRange();
	range.setStartAfter(heading);
	let endNode = heading.parentElement;
	let next = heading.nextElementSibling;
	while (next) {
		if (/^H[1-6]$/.test(next.tagName)) {
			const nextLevel = Number(next.tagName.substring(1));
			if (nextLevel <= level) break; // 到达同级或更高层级，范围结束
		}
		endNode = next;
		next = next.nextElementSibling;
	}
	if (endNode) {
		range.setEndAfter(endNode);
	} else {
		range.setEndAfter(heading.parentElement.lastChild);
	}
	const wrapper = document.createElement('div');
	range.surroundContents(wrapper);
	// 复制节点，避免改变 DOM 结构（我们只用来定位范围，不保留包装）
	const clone = wrapper.cloneNode(true);
	// 还原
	const parent = wrapper.parentNode;
	while (wrapper.firstChild) parent.insertBefore(wrapper.firstChild, wrapper);
	parent.removeChild(wrapper);
	return clone;
}

// 加载静态 content.html 到正文（若存在）
async function loadStaticContent() {
	const article = document.getElementById('article');
	try {
		const res = await fetch('./content.html', { cache: 'no-store' });
		if (!res.ok) return; // 文件不存在则跳过，继续使用内置示例内容
		const html = await res.text();
		article.innerHTML = html;
		// 为没有 id 的标题补 id，便于目录
		article.querySelectorAll('h1,h2,h3').forEach((h) => {
			if (!h.id) {
				h.id = h.textContent.trim().replace(/\s+/g, '-').toLowerCase();
			}
		});
		// 重建目录与监听、范围
		buildToc();
		enableTocClick();
		observeHeadings();
		// 清除旧高亮（如果有）
		clearHighlights(article);
		// 回到顶部
		document.querySelector('.content').scrollTop = 0;
	} catch (_) {
		// 忽略加载失败，保留默认示例内容
	}
}

// 文章列表与加载
async function loadArticleList() {
	const listNav = document.getElementById('article-list');
	if (!listNav) return false;
	try {
		const res = await fetch('./articles/articles.json', { cache: 'no-store' });
		if (!res.ok) return false;
		/** @type {{title:string,file:string,id?:string}[]} */
		const articles = await res.json();
		if (!Array.isArray(articles) || articles.length === 0) return false;

		listNav.innerHTML = '';
		const ul = document.createElement('ul');
		articles.forEach((item, idx) => {
			const li = document.createElement('li');
			li.className = 'lvl-1';
			const a = document.createElement('a');
			a.href = '#';
			a.textContent = item.title || `文章 ${idx + 1}`;
			a.dataset.file = item.file;
			a.addEventListener('click', (e) => {
				e.preventDefault();
				loadArticleFile(a.dataset.file, listNav, li);
			});
			li.appendChild(a);
			ul.appendChild(li);
		});
		listNav.appendChild(ul);

		// 默认加载第一篇
		const first = ul.querySelector('li a');
		if (first && first.dataset.file) {
			loadArticleFile(first.dataset.file, listNav, first.parentElement);
		}
		return true;
	} catch (_) {
		return false;
	}
}

async function loadArticleFile(path, listNav, liEl) {
	const article = document.getElementById('article');
	try {
		// .docx 走 mammoth 转 HTML，.html 按文本加载
		if (/\.docx$/i.test(path)) {
			const res = await fetch(path, { cache: 'no-store' });
			if (!res.ok) throw new Error('无法获取 .docx 文件');
			const arrayBuffer = await res.arrayBuffer();
			const result = await window.mammoth.convertToHtml({ arrayBuffer });
			article.innerHTML = result.value || '<p>未解析到内容</p>';
		} else {
			const res = await fetch(path, { cache: 'no-store' });
			if (!res.ok) {
				alert('文章加载失败：' + path);
				return;
			}
			const html = await res.text();
			article.innerHTML = html;
		}
		// 标题补 id
		article.querySelectorAll('h1,h2,h3').forEach((h) => {
			if (!h.id) h.id = h.textContent.trim().replace(/\s+/g, '-').toLowerCase();
		});
		// 高亮当前选中文章
		listNav.querySelectorAll('li').forEach((li) => li.classList.remove('active'));
		if (liEl) liEl.classList.add('active');
		// 重建章节目录与监听
		buildToc();
		enableTocClick();
		observeHeadings();
		// 清除旧检索高亮与计数
		clearHighlights(article);
		document.getElementById('search-count').textContent = '0';
		document.querySelector('.content').scrollTop = 0;
	} catch (e) {
		console.error(e);
		alert('文章加载失败：' + path);
	}
}

// DOCX -> index.html + images 导出为 Zip
function setupDocxExtractor() {
	const input = document.getElementById('docx-extract-input');
	const btn = document.getElementById('extract-btn');
	if (!input || !btn) return;

	function mimeToExt(mime) {
		if (!mime) return 'png';
		if (mime === 'image/png') return 'png';
		if (mime === 'image/jpeg' || mime === 'image/jpg') return 'jpg';
		if (mime === 'image/gif') return 'gif';
		if (mime === 'image/svg+xml') return 'svg';
		if (mime === 'image/bmp') return 'bmp';
		return 'png';
	}

	btn.addEventListener('click', async () => {
		const file = input.files && input.files[0];
		if (!file) {
			alert('请先选择一个 .docx 文件');
			return;
		}
		if (!/\.docx$/i.test(file.name)) {
			alert('仅支持 .docx 文件');
			return;
		}

		const zip = new JSZip();
		const imagesFolder = zip.folder('images');
		let imgIndex = 0;

		try {
			const arrayBuffer = await file.arrayBuffer();
			const result = await window.mammoth.convertToHtml(
				{ arrayBuffer },
				{
					convertImage: window.mammoth.images.inline(async (element) => {
						// 读取为 base64 并写入 zip 的 images/ 目录，返回相对路径
						const base64 = await element.read('base64');
						const ext = mimeToExt(element.contentType);
						const filename = `img${String(++imgIndex).padStart(3, '0')}.${ext}`;
						if (imagesFolder) {
							imagesFolder.file(filename, base64, { base64: true });
						}
						return { src: `images/${filename}` };
					})
				}
			);

			// 包装为可直接打开的 HTML 文件
			const title = file.name.replace(/\.docx$/i, '');
			const htmlDoc =
				'<!doctype html>\n' +
				'<html lang="zh-CN">\n' +
				'<head>\n' +
				'\t<meta charset="utf-8">\n' +
				`\t<title>${title}</title>\n` +
				'\t<meta name="viewport" content="width=device-width, initial-scale=1">\n' +
				'\t<style>\n' +
				'\t\tbody{max-width:900px;margin:24px auto;font-family:system-ui,-apple-system,"Segoe UI",Roboto,"Helvetica Neue",Arial,"Noto Sans","PingFang SC","Microsoft Yahei",sans-serif;line-height:1.75;color:#111827}\n' +
				'\t\th1{font-size:28px;margin:0 0 16px}\n' +
				'\t\th2{font-size:22px;margin-top:28px;color:#1d4ed8}\n' +
				'\t\th3{font-size:18px;margin-top:20px;color:#2563eb}\n' +
				'\t\tp{color:#374151}\n' +
				'\t\timg{max-width:100%;height:auto}\n' +
				'\t</style>\n' +
				'</head>\n' +
				'<body>\n' +
				(result.value || '<p>未解析到内容</p>') +
				'\n</body>\n</html>\n';

			zip.file('index.html', htmlDoc);

			const blob = await zip.generateAsync({ type: 'blob' });
			const zipName = `${title}.zip`;
			saveAs(blob, zipName);
		} catch (err) {
			console.error(err);
			alert('提取失败：请确认文件为有效的 .docx');
		}
	});
}

// 执行检索
function setupSearch() {
	const input = document.getElementById('search-input');
	const btn = document.getElementById('search-btn');
	const clearBtn = document.getElementById('clear-btn');
	const countEl = document.getElementById('search-count');
	const scopeSelect = document.getElementById('scope-select');
	const article = document.getElementById('article');

	function runSearch() {
		// 清旧高亮
		clearHighlights(article);
		countEl.textContent = '0';
		const q = input.value.trim();
		if (!q) return;

		let total = 0;
		const scope = scopeSelect.value;
		if (scope === '__ALL__') {
			total = highlightText(article, q);
		} else {
			// 仅在选中标题的同级范围内高亮
			const heading = document.getElementById(scope);
			if (heading) {
				// 计算范围：用 Range 克隆内容，在克隆上计算不中断真实 DOM，再在真实范围内高亮
				// 为精确高亮真实 DOM，我们需要定位真实范围的起止节点集合
				// 简化：在真实 DOM 中找到从 heading.nextSibling 开始到下一个同级标题前的节点，逐个容器内高亮
				const level = Number(heading.tagName.substring(1));
				let node = heading.nextSibling;
				let stopAt = null;
				let iter = heading.nextElementSibling;
				while (iter) {
					if (/^H[1-6]$/.test(iter.tagName) && Number(iter.tagName.substring(1)) <= level) {
						stopAt = iter;
						break;
					}
					iter = iter.nextElementSibling;
				}
				const containers = [];
				while (node && node !== stopAt) {
					if (node.nodeType === Node.ELEMENT_NODE) {
						containers.push(node);
					}
					node = node.nextSibling;
				}
				containers.forEach((el) => {
					total += highlightText(el, q);
				});
			}
		}
		countEl.textContent = String(total);
	}

	btn.addEventListener('click', runSearch);
	input.addEventListener('keydown', (e) => {
		if (e.key === 'Enter') runSearch();
	});
	clearBtn.addEventListener('click', () => {
		clearHighlights(article);
		countEl.textContent = '0';
	});
}

function init() {
	buildToc();
	enableTocClick();
	observeHeadings();
	// 优先加载文章列表；若不存在则加载单页 content.html
	loadArticleList().then((loaded) => {
		if (!loaded) {
			loadStaticContent();
		}
	});
	setupDocxExtractor();
	setupSearch();
}

document.addEventListener('DOMContentLoaded', init);


