const bannedPhrases = {
  "Opinion or weak certainty": [
    "i think",
    "in my opinion",
    "pretty sure",
    "i am sure",
    "i'm sure",
    "it might",
    "it may",
    "perhaps",
    "probably",
    "possibly",
  ],
  "Back-reference to earlier content": [
    "as stated before",
    "as mentioned above",
    "as discussed earlier",
    "as explained in",
    "as it is explained",
    "as noted earlier",
  ],
  "Everyday or filler language": [
    "basically",
    "kind of",
    "sort of",
    "a lot of",
    "you know",
    "stuff",
    "things",
    "really",
    "very",
  ],
  "Analogy markers": ["think of it as", "imagine", "just like", "similar to a", "as if"],
  "Stressful or hateful wording": ["you will be dead", "hate", "hateful", "awful", "terrible", "disaster", "catastrophe"],
}

const guidelineSummary = [
  "Use AP-style discipline, factual scientific prose, and active voice.",
  "Avoid opinions, analogies, everyday language, weak certainty, and contextless words.",
  "Do not start sentences with if or because, and do not use the word also.",
  "Use one focus, keep the central entity in every section, and answer search intent in the intro.",
  "Write short, complete sentences that follow subject-predicate-object declarations.",
  "Answer question headings directly, with certainty, and without distancing the answer.",
  "Use facts, consensus, research details, examples, data points, units, and percentages.",
  "Introduce abbreviations at first mention, such as Bitcoin (BTC).",
  "Keep featured snippets within 40 words and 320 characters.",
  "Avoid first-paragraph anchors and anchor text as the first words of a paragraph.",
]

const modalWords = new Set(["can", "could", "may", "might", "should", "would"])
const researchWords = ["study", "research", "researcher", "journal", "publication", "survey", "analysis", "data", "dataset", "consensus", "according to", "reported", "observed"]
const wordPattern = /\b[A-Za-z][A-Za-z'-]*\b/g

function normalizeText(text) {
  return String(text || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim()
}

function words(text) {
  return text.match(wordPattern) || []
}

function wordCount(text) {
  return words(text).length
}

function splitSentences(text) {
  return text
    .replace(/\s+/g, " ")
    .split(/(?<=[.!?])\s+/)
    .map((part) => part.trim())
    .filter((part) => wordPattern.test(part) && (wordPattern.lastIndex = 0) === 0)
}

function splitParagraphs(text) {
  return text.split(/\n\s*\n/).map((part) => part.trim()).filter(Boolean)
}

function excerpt(value, limit = 220) {
  const clean = String(value || "").replace(/\s+/g, " ").trim()
  return clean.length <= limit ? clean : `${clean.slice(0, limit - 1).trimEnd()}...`
}

function lineNumberFor(text, snippet) {
  if (!snippet) return null
  const index = text.toLowerCase().indexOf(snippet.toLowerCase().slice(0, 80))
  return index < 0 ? null : text.slice(0, index).split("\n").length
}

function addIssue(issues, text, severity, category, rule, message, suggestion, evidence) {
  issues.push({ severity, category, rule, message, suggestion, evidence: excerpt(evidence), line: lineNumberFor(text, evidence) })
}

function inferSections(text) {
  const sections = []
  let currentTitle = "Introduction"
  let currentBody = []
  const looksLikeHeading = (line) => {
    const stripped = line.trim().replace(/^#+/, "").trim()
    if (!stripped || stripped.length > 90 || stripped.endsWith(".")) return false
    const count = stripped.split(/\s+/).length
    return count <= 9 && (line.trimStart().startsWith("#") || stripped.endsWith("?") || stripped.endsWith(":") || stripped === stripped.replace(/\b\w/g, (c) => c.toUpperCase()))
  }
  for (const raw of text.split("\n")) {
    const line = raw.trim()
    if (looksLikeHeading(line)) {
      if (currentBody.length) sections.push([currentTitle, currentBody.join("\n")])
      currentTitle = line.replace(/^#+\s*/, "").replace(/:$/, "")
      currentBody = []
    } else {
      currentBody.push(raw)
    }
  }
  if (currentBody.length) sections.push([currentTitle, currentBody.join("\n")])
  return sections.map(([title, body]) => [title, normalizeText(body)]).filter(([, body]) => body)
}

function auditContent(rawText, centralEntity = "", searchIntent = "") {
  const text = normalizeText(rawText)
  if (!text) throw new Error("Add pasted text or upload a content file before running the audit.")

  const sentences = splitSentences(text)
  const paragraphs = splitParagraphs(text)
  const sections = inferSections(text)
  const issues = []
  const lowered = text.toLowerCase()

  for (const [category, phrases] of Object.entries(bannedPhrases)) {
    for (const phrase of phrases) {
      const index = lowered.search(new RegExp(`\\b${phrase.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`))
      if (index >= 0) {
        addIssue(issues, text, ["Opinion or weak certainty", "Back-reference to earlier content"].includes(category) ? "high" : "medium", "Tone and wording", category, `Guidelines discourage '${phrase}'.`, "Replace it with factual, direct, context-specific wording.", text.slice(index, index + 180))
      }
    }
  }

  const also = lowered.search(/\balso\b/)
  if (also >= 0) addIssue(issues, text, "high", "Tone and wording", "Banned word", "The guideline says never use the word 'also'.", "Remove it or replace it with a more precise transition.", text.slice(also, also + 160))

  for (const sentence of sentences) {
    const count = wordCount(sentence)
    if (count > 25) addIssue(issues, text, "medium", "Sentence quality", "Short sentences", `This sentence has ${count} words.`, "Split it so each sentence gives one clear piece of information.", sentence)
    const first = words(sentence)[0]?.toLowerCase()
    if (first === "if" || first === "because") addIssue(issues, text, "high", "Sentence quality", "Opening conditional", `The sentence starts with '${words(sentence)[0]}'.`, "Move the conditional or cause clause to the second part of the sentence.", sentence)
    if (/\b(?:is|are|was|were|be|been|being|has been|have been|had been)\s+\w+(?:ed|en)\b/i.test(sentence)) addIssue(issues, text, "medium", "Sentence quality", "Active voice", "This sentence appears to use passive voice.", "Rewrite with a clear subject performing the action.", sentence)
    const foundModals = [...new Set(words(sentence).map((w) => w.toLowerCase()).filter((w) => modalWords.has(w)))]
    if (foundModals.length) addIssue(issues, text, "medium", "Certainty", "Possibility modal", `Modal wording reduces certainty: ${foundModals.sort().join(", ")}.`, "Use definitive wording unless the heading question uses the same modality.", sentence)
  }

  const entity = centralEntity.trim()
  const intro = paragraphs.slice(0, 2).join("\n\n")
  if (entity) {
    const entityRe = new RegExp(`\\b${entity.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`, "i")
    if (!entityRe.test(intro)) addIssue(issues, text, "high", "Semantic focus", "Central entity in intro", "The central entity is missing from the intro.", "Mention the central entity early and process the article through that entity.", intro || text.slice(0, 220))
    const missing = sections.filter(([title, body]) => title !== "Introduction" && !entityRe.test(`${title} ${body}`)).map(([title]) => title)
    if (missing.length) addIssue(issues, text, "high", "Semantic focus", "Central entity in sections", "Some sections do not mention the central entity.", "Use the central entity or a clear synonym in every section.", missing.slice(0, 6).join(", "))
  }

  const intentTerms = words(searchIntent).map((w) => w.toLowerCase()).filter((w) => w.length > 3)
  const introTerms = new Set(words(intro).map((w) => w.toLowerCase()))
  if (intentTerms.some((term) => !introTerms.has(term))) addIssue(issues, text, "medium", "Semantic focus", "Search intent in intro", "The intro does not clearly cover the central search intent.", "Answer the core user intent in the intro before expanding details.", intro || text.slice(0, 220))

  const researchMentions = researchWords.reduce((total, term) => total + (lowered.match(new RegExp(`\\b${term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`, "g")) || []).length, 0)
  const years = (text.match(/\b(?:19|20)\d{2}\b/g) || []).length
  const numbers = (text.match(/\b\d+(?:[.,]\d+)?\b/g) || []).length
  const units = (text.match(/\b(?:kg|g|mg|lb|lbs|oz|km|m|cm|mm|mi|ft|in|mph|km\/h|kwh|kw|w|v|hz|gb|mb|ms|s|sec|mins?|hours?|days?|years?|usd|rs|lkr|%|percent)\b/gi) || []).length
  const percentages = (text.match(/\b\d+(?:\.\d+)?\s*(?:%|percent)\b/gi) || []).length
  const examples = (text.match(/\b(?:for example|for instance|such as|e\.g\.)\b/gi) || []).length

  if (sentences.length >= 8 && researchMentions === 0) addIssue(issues, text, "medium", "Evidence", "Research support", "No research, study, data, or consensus terms were found.", "Integrate research source, organization, date, topic, sample count, or consensus details.", text.slice(0, 220))
  if (sentences.length >= 8 && numbers < 2) addIssue(issues, text, "low", "Evidence", "Numeric detail", "The article has limited numeric detail.", "Add precise values, counts, ranges, dates, units, or percentages where relevant.", text.slice(0, 220))
  if (sentences.length >= 8 && examples < 2) addIssue(issues, text, "low", "Evidence", "Examples", "The article has limited example language.", "Use multiple examples to create specific contextual connections.", text.slice(0, 220))

  const severityRank = { high: 0, medium: 1, low: 2 }
  issues.sort((a, b) => severityRank[a.severity] - severityRank[b.severity] || a.category.localeCompare(b.category) || (a.line || 999999) - (b.line || 999999))
  const high = issues.filter((issue) => issue.severity === "high").length
  const medium = issues.filter((issue) => issue.severity === "medium").length
  const low = issues.filter((issue) => issue.severity === "low").length
  const sentenceLengths = sentences.map(wordCount)
  const average = sentenceLengths.length ? Math.round((sentenceLengths.reduce((a, b) => a + b, 0) / sentenceLengths.length) * 10) / 10 : 0
  const score = Math.max(0, Math.min(100, 100 - high * 12 - medium * 7 - low * 3))

  return {
    score,
    summary: { words: wordCount(text), sentences: sentences.length, paragraphs: paragraphs.length, sections: sections.length, average_sentence_words: average, high, medium, low },
    detail_metrics: { research_mentions: researchMentions, years, numbers, units, percentages, examples },
    guidelines: guidelineSummary,
    issues,
  }
}

export default async (req) => {
  if (req.method !== "POST") return Response.json({ error: "Method not allowed." }, { status: 405 })

  try {
    const form = await req.formData()
    const pastedText = form.get("contentText")?.toString() || ""
    const centralEntity = form.get("centralEntity")?.toString() || ""
    const searchIntent = form.get("searchIntent")?.toString() || ""
    const file = form.get("contentFile")
    let uploadedText = ""

    if (file && typeof file.text === "function" && file.size > 0) {
      const name = file.name || ""
      if (!/\.(txt|md|markdown|html|htm|csv|rtf)$/i.test(name)) {
        return Response.json({ error: "This deployed version supports text-based uploads: .txt, .md, .html, .csv, or .rtf. Paste PDF or DOCX text into the content field." }, { status: 400 })
      }
      uploadedText = await file.text()
    }

    return Response.json(auditContent([uploadedText, pastedText].filter((part) => part.trim()).join("\n\n"), centralEntity, searchIntent))
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unexpected audit error."
    return Response.json({ error: message }, { status: message.startsWith("Add pasted") ? 400 : 500 })
  }
}

export const config = {
  path: "/audit",
}
