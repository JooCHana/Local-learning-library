from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from library_core import LibraryDatabase, SearchHit


GENERIC_QUESTION_WORDS = (
    "请根据", "根据文档", "根据资料", "在文档库中", "文档库内", "学习库内", "请问", "请查询",
    "请说明", "请简述", "请列出", "查找", "查询", "检索", "整理", "回答", "什么", "哪些",
    "如何", "怎样", "怎么", "是否", "有关", "相关", "要求", "规定", "内容", "问题",
)

OPPOSITE_GROUPS = (
    (("允许", "可以", "准许", "可办理"), ("禁止", "不得", "不准", "严禁", "不可")),
    (("必须", "应当", "应予"), ("无需", "不必", "可以不")),
    (("是", "属于"), ("不是", "不属于")),
)


@dataclass(frozen=True)
class SourceReference:
    doc_id: int
    name: str
    path: str
    location: str
    document_number: str = ""


@dataclass
class OrganizedStatement:
    text: str
    score: float
    sources: list[SourceReference] = field(default_factory=list)
    conflict_with: list[int] = field(default_factory=list)


@dataclass
class QuestionResearchResult:
    question: str
    keywords: list[str]
    statements: list[OrganizedStatement]
    documents: list[SourceReference]
    matched_chunks: int

    def render_text(self) -> str:
        lines = [
            f"问题：{self.question}",
            f"核心关键词：{'、'.join(self.keywords) if self.keywords else '未提取到'}",
            f"检索到 {self.matched_chunks} 条原文片段，去重整理为 {len(self.statements)} 条唯一表述。",
            "",
            "整理结果（以下内容均来自文档原文）",
        ]
        if not self.statements:
            lines.extend(["", "未在当前检索范围内找到足够匹配的文档原文。"])
        for index, statement in enumerate(self.statements, 1):
            lines.extend(["", f"{index}. {statement.text}"])
            source_text = "；".join(
                f"{source.name} · {source.location}" for source in statement.sources
            )
            lines.append(f"   来源：{source_text}")
            if statement.conflict_with:
                related = "、".join(str(value) for value in statement.conflict_with)
                lines.append(f"   ⚠ 与第 {related} 条可能存在表述冲突或版本差异，请结合文号和施行时间核对。")

        lines.extend(["", "使用的文档名称 / 文号"])
        if not self.documents:
            lines.append("- 无")
        else:
            for source in self.documents:
                number = source.document_number or "未识别到文号"
                lines.append(f"- {source.name} / {number}")
        return "\n".join(lines)

    def render_markdown(self) -> str:
        lines = [
            "# 文档问题检索整理",
            "",
            f"**问题：** {self.question}",
            "",
            f"**核心关键词：** {'、'.join(self.keywords) if self.keywords else '未提取到'}",
            "",
            f"> 检索到 {self.matched_chunks} 条原文片段，去重整理为 {len(self.statements)} 条唯一表述。以下内容均来自文档原文。",
            "",
            "## 整理结果",
            "",
        ]
        if not self.statements:
            lines.append("未在当前检索范围内找到足够匹配的文档原文。")
        for index, statement in enumerate(self.statements, 1):
            lines.extend([f"{index}. {statement.text}", ""])
            lines.append("   来源：" + "；".join(
                f"{source.name} · {source.location}" for source in statement.sources
            ))
            if statement.conflict_with:
                related = "、".join(str(value) for value in statement.conflict_with)
                lines.append(f"   **注意：** 与第 {related} 条可能存在表述冲突或版本差异，请结合文号和施行时间核对。")
            lines.append("")
        lines.extend(["## 使用的文档名称 / 文号", ""])
        if not self.documents:
            lines.append("- 无")
        else:
            for source in self.documents:
                lines.append(f"- {source.name} / {source.document_number or '未识别到文号'}")
        return "\n".join(lines)


def extract_core_keywords(question: str, maximum: int = 18) -> list[str]:
    """Extract deterministic search phrases without inventing domain concepts."""
    question = unicodedata.normalize("NFKC", question).strip()
    if not question:
        return []
    candidates: list[tuple[int, str]] = []

    def add(value: str, priority: int = 0) -> None:
        value = value.strip(" ，。；：、？！?!.;:()（）[]【】\"'“”")
        if len(value) >= 2:
            candidates.append((priority + min(len(value), 12), value))

    for quoted in re.findall(r"《([^》]{2,40})》|[“\"]([^”\"]{2,40})[”\"]", question):
        add(next(item for item in quoted if item), 40)
    for reference in re.findall(
        r"[\u4e00-\u9fff]{0,10}[〔\[【(]\d{4}[〕\]】)]\s*\d+\s*号|\b[A-Za-z]{2,}[\w./-]*\b|\d+(?:\.\d+)?(?:%|‰|公里|米|吨|分钟|小时|日)",
        question,
    ):
        add(reference, 35)

    cleaned = question
    for word in GENERIC_QUESTION_WORDS:
        cleaned = cleaned.replace(word, " ")
    cleaned = re.sub(r"[，。；：、？！?!.;:\n\r\t()（）\[\]【】{}“”\"']+", " ", cleaned)
    for token in cleaned.split():
        if re.fullmatch(r"[A-Za-z0-9_./%+-]{2,}", token):
            add(token, 25)
            continue
        for run in re.findall(r"[\u3400-\u9fff]{2,}", token):
            if len(run) <= 12:
                add(run, 28)
            # Overlapping longer phrases retrieve Chinese provisions more
            # precisely than isolated two-character words.
            for width, priority in ((8, 22), (6, 18), (4, 10)):
                if len(run) <= width:
                    continue
                step = max(2, width // 2)
                for start in range(0, len(run) - width + 1, step):
                    add(run[start : start + width], priority)
                add(run[-width:], priority)

    unique: dict[str, int] = {}
    for priority, value in candidates:
        key = value.casefold()
        unique[key] = max(priority, unique.get(key, -1))
    ordered = sorted(unique, key=lambda key: (-unique[key], -len(key), key))
    return [next(value for _score, value in candidates if value.casefold() == key) for key in ordered[:maximum]]


def _canonical_text(text: str) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", text)).strip()


def _split_source_statements(content: str) -> list[str]:
    statements: list[str] = []
    for paragraph in re.split(r"\n\s*\n+", content):
        paragraph = re.sub(r"[ \t]+", " ", paragraph).strip()
        if not paragraph:
            continue
        # PDF layout extraction often inserts line breaks inside one sentence.
        paragraph = re.sub(r"(?<![。！？；：])\n(?!\s*$)", "", paragraph)
        paragraph = re.sub(r"\n+", "\n", paragraph)
        parts = [item.strip() for item in re.findall(r"[^。！？；\n]+[。！？；]?", paragraph) if item.strip()]
        if not parts:
            parts = [paragraph]
        pending_heading = ""
        for part in parts:
            if re.fullmatch(r"第[一二三四五六七八九十百千万零〇0-9]+条[:：]?", part):
                pending_heading = part
                continue
            if pending_heading:
                part = pending_heading + part
                pending_heading = ""
            if len(_canonical_text(part)) >= 6:
                statements.append(part)
        if pending_heading:
            statements.append(pending_heading)
    return statements


def _keyword_score(text: str, keywords: list[str]) -> float:
    canonical = _canonical_text(text).casefold()
    matched = [keyword for keyword in keywords if _canonical_text(keyword).casefold() in canonical]
    if not matched:
        return 0.0
    return sum(len(keyword) ** 1.6 for keyword in matched) + len(matched) * 4


def _char_ngrams(text: str, width: int = 2) -> set[str]:
    clean = re.sub(r"[^\u3400-\u9fffA-Za-z0-9]", "", text).casefold()
    return {clean[index : index + width] for index in range(max(0, len(clean) - width + 1))}


def _possibly_conflicting(left: str, right: str) -> bool:
    left_grams, right_grams = _char_ngrams(left), _char_ngrams(right)
    if not left_grams or not right_grams:
        return False
    similarity = len(left_grams & right_grams) / max(1, len(left_grams | right_grams))
    same_article = bool(
        set(re.findall(r"第[一二三四五六七八九十百千万零〇0-9]+条", left))
        & set(re.findall(r"第[一二三四五六七八九十百千万零〇0-9]+条", right))
    )
    if similarity < 0.34 and not same_article:
        return False
    for positive, negative in OPPOSITE_GROUPS:
        left_positive = any(word in left for word in positive)
        left_negative = any(word in left for word in negative)
        right_positive = any(word in right for word in positive)
        right_negative = any(word in right for word in negative)
        if (left_positive and right_negative) or (left_negative and right_positive):
            return True
    return same_article and 0.50 <= similarity < 0.92


def extract_document_number(name: str, content: str = "") -> str:
    sample = f"{name} {content[:1200]}"
    patterns = (
        r"[\u4e00-\u9fff]{1,12}[〔\[【(]\s*\d{4}\s*[〕\]】)]\s*\d+\s*号",
        r"[\u4e00-\u9fff]{0,8}\d{2,4}\s*号",
    )
    found: list[str] = []
    for pattern in patterns:
        for value in re.findall(pattern, sample):
            value = re.sub(r"\s+", "", value)
            if value not in found:
                found.append(value)
    return "、".join(found[:3])


def organize_question_from_library(
    database: LibraryDatabase,
    question: str,
    categories: str | Iterable[str] = "",
) -> QuestionResearchResult:
    question = question.strip()
    if not question:
        raise ValueError("请输入需要检索整理的问题")
    keywords = extract_core_keywords(question)
    if not keywords:
        raise ValueError("未能从问题中提取有效关键词，请补充具体名词、文号、数字或条款内容")

    hit_scores: dict[int, tuple[SearchHit, float]] = {}
    for rank, keyword in enumerate(keywords):
        weight = max(1.0, len(keywords) - rank) * max(2, len(keyword))
        # Use a deliberately high ceiling so this feature compares the full
        # library result set instead of the paged result list shown on home.
        for hit in database.search(keyword, category=categories, limit=100000):
            current = hit_scores.get(hit.chunk_id)
            amount = weight + hit.score * len(keyword)
            if current:
                hit_scores[hit.chunk_id] = (hit, current[1] + amount)
            else:
                hit_scores[hit.chunk_id] = (hit, amount)

    deduplicated: dict[str, OrganizedStatement] = {}
    first_order: dict[str, int] = {}
    document_content: dict[int, tuple[SearchHit, str]] = {}
    ordered_hits = sorted(hit_scores.values(), key=lambda item: (-item[1], item[0].name, item[0].location))
    for order, (hit, chunk_score) in enumerate(ordered_hits):
        document_content.setdefault(hit.doc_id, (hit, hit.content))
        source = SourceReference(
            doc_id=hit.doc_id,
            name=hit.name,
            path=hit.path,
            location=hit.location,
            document_number=extract_document_number(hit.name, hit.content),
        )
        for statement_text in _split_source_statements(hit.content):
            relevance = _keyword_score(statement_text, keywords)
            if relevance <= 0:
                continue
            key = _canonical_text(statement_text).casefold()
            if not key:
                continue
            statement = deduplicated.get(key)
            if statement is None:
                statement = OrganizedStatement(statement_text, relevance + chunk_score * 0.05, [source])
                deduplicated[key] = statement
                first_order[key] = order
            else:
                statement.score = max(statement.score, relevance + chunk_score * 0.05)
                if source not in statement.sources:
                    statement.sources.append(source)

    statements = sorted(
        deduplicated.values(),
        key=lambda item: (-item.score, first_order[_canonical_text(item.text).casefold()], item.text),
    )
    # Conflict comparison does not rewrite or discard either source statement.
    # It only flags pairs sharing a subject/article while using opposing wording.
    for left_index, left in enumerate(statements):
        for right_index in range(left_index + 1, len(statements)):
            right = statements[right_index]
            if _possibly_conflicting(left.text, right.text):
                left.conflict_with.append(right_index + 1)
                right.conflict_with.append(left_index + 1)

    used_doc_ids = sorted({source.doc_id for item in statements for source in item.sources})
    documents: list[SourceReference] = []
    if used_doc_ids:
        placeholders = ",".join("?" for _ in used_doc_ids)
        with database.connect() as connection:
            rows = connection.execute(
                f"""SELECT id, name, path, full_text FROM documents
                    WHERE id IN ({placeholders}) ORDER BY name COLLATE NOCASE""",
                used_doc_ids,
            ).fetchall()
        for row in rows:
            documents.append(
                SourceReference(
                    doc_id=int(row["id"]),
                    name=str(row["name"]),
                    path=str(row["path"]),
                    location="",
                    document_number=extract_document_number(str(row["name"]), str(row["full_text"])),
                )
            )
    return QuestionResearchResult(question, keywords, statements, documents, len(hit_scores))


def export_question_research_markdown(result: QuestionResearchResult, target: str | Path) -> Path:
    target = Path(target)
    target.write_text(result.render_markdown(), encoding="utf-8-sig")
    return target


def export_question_research_word(result: QuestionResearchResult, target: str | Path) -> Path:
    target = Path(target)
    document = Document()
    title = document.add_paragraph()
    title_run = title.add_run("文档问题检索整理")
    title_run.bold = True
    title_run.font.size = Pt(16)
    title.paragraph_format.space_after = Pt(12)
    document.add_paragraph(f"问题：{result.question}")
    document.add_paragraph(f"核心关键词：{'、'.join(result.keywords) if result.keywords else '未提取到'}")
    document.add_paragraph(
        f"检索到 {result.matched_chunks} 条原文片段，去重整理为 {len(result.statements)} 条唯一表述。"
        "以下内容均来自文档原文。"
    )
    document.add_heading("整理结果", level=1)
    if not result.statements:
        document.add_paragraph("未在当前检索范围内找到足够匹配的文档原文。")
    for index, statement in enumerate(result.statements, 1):
        document.add_paragraph(statement.text, style="List Number")
        source = document.add_paragraph(
            "来源：" + "；".join(f"{item.name} · {item.location}" for item in statement.sources)
        )
        if source.runs:
            source.runs[0].italic = True
        if statement.conflict_with:
            related = "、".join(str(value) for value in statement.conflict_with)
            warning = document.add_paragraph(
                f"注意：与第 {related} 条可能存在表述冲突或版本差异，请结合文号和施行时间核对。"
            )
            for run in warning.runs:
                run.bold = True
    document.add_heading("使用的文档名称 / 文号", level=1)
    if not result.documents:
        document.add_paragraph("无")
    else:
        for source in result.documents:
            document.add_paragraph(
                f"{source.name} / {source.document_number or '未识别到文号'}", style="List Bullet"
            )

    style_sizes = {"Normal": 10.5, "Heading 1": 12}
    for style_name, size in style_sizes.items():
        style = document.styles[style_name]
        style.font.name = "Times New Roman"
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor(0, 0, 0)
        style._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "宋体")
    for paragraph in document.paragraphs:
        paragraph.paragraph_format.line_spacing = 1.0
        size = 16 if paragraph._p is title._p else style_sizes.get(paragraph.style.name, 10.5)
        for run in paragraph.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(size)
            run.font.color.rgb = RGBColor(0, 0, 0)
            run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "宋体")
    document.save(target)
    return target
