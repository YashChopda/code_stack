const maxArticles = $node["Input Config"].json.maxArticles || 10;

const entries = $json.feed?.entry || [];
const sliced = entries.slice(0, maxArticles);

return sliced.map(e => ({
  json: {
    title: e.title || "",
    url: Array.isArray(e.link) ? (e.link.find(l => l.href)?.href || "") : (e.link?.href || ""),
    updated: e.updated || "",
    author: e.author?.name || "",
    contentHtml: e.content || "",
    id: e.id || ""
  }
}));
