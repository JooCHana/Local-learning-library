from pathlib import Path
from tempfile import TemporaryDirectory

from library_core import LibraryDatabase, extract_docx, extract_markdown, extract_pdf


ROOT = Path(__file__).resolve().parent.parent


def main() -> None:
    markdown = Path(__file__).with_name("使用说明.md")
    docx = ROOT / "题库截图识别助手" / "release" / "示例输出.docx"
    pdf = ROOT / "题库截图识别助手" / "test_output" / "office_render" / "word.pdf"

    assert extract_markdown(markdown), "Markdown 提取失败"
    assert extract_docx(docx), "Word 提取失败"
    assert extract_pdf(pdf), "PDF 提取失败"

    with TemporaryDirectory() as temp:
        temp_root = Path(temp)
        database = LibraryDatabase(temp_root / "test.db", temp_root / "资料文件")
        status, count = database.import_file(markdown, category="铁路规章")
        assert status == "added" and count > 0
        assert database.search("关键词"), "中文关键词搜索失败"
        assert database.import_file(markdown, category="铁路规章")[0] == "duplicate", "重复文件未识别"
        rows = database.list_documents()
        assert len(rows) == 1
        doc_id = int(rows[0]["id"])
        assert database.search("关键词", doc_id=doc_id), "单文件限定搜索失败"
        assert not database.search("关键词", doc_id=doc_id + 9999), "单文件搜索范围未生效"
        assert database.find_documents("铁路规章"), "按分类查找文档失败"
        assert database.list_documents(sort_by="date_desc"), "按添加日期排序失败"
        stored = Path(rows[0]["stored_path"])
        assert stored.exists() and stored.parent.name == "铁路规章", "原始文件未复制到分类文件夹"
        stored_mtime = stored.stat().st_mtime_ns
        overwrite_status, _ = database.import_file(
            markdown, duplicate_action="overwrite", category="铁路规章"
        )
        assert overwrite_status == "skipped", "相同内容覆盖导入应直接复用已有副本"
        assert stored.stat().st_mtime_ns == stored_mtime, "相同内容不应重复改写学习库副本"
        database.update_tags(rows[0]["id"], "复习,铁路")
        database.update_note(rows[0]["id"], "重点复习资料")
        assert database.search("关键词", tag="复习"), "标签筛选失败"
        assert database.search("关键词", category="铁路规章"), "分类筛选失败"
        database.move_document_category(rows[0]["id"], "重点资料")
        rows = database.list_documents()
        stored = Path(rows[0]["stored_path"])
        assert rows[0]["category"] == "重点资料" and stored.parent.name == "重点资料", "分类移动失败"
        backup = temp_root / "backup.zip"
        database.export_backup(backup)
        assert backup.exists(), "备份导出失败"
        assert database.delete_documents([rows[0]["id"]]) == 1
        assert not stored.exists(), "删除记录时未删除资料库文件副本"
        assert database.counts() == (0, 0)
        database.restore_backup(backup)
        restored = database.list_documents()
        assert (
            len(restored) == 1
            and restored[0]["tags"] == "复习,铁路"
            and restored[0]["category"] == "重点资料"
            and restored[0]["note"] == "重点复习资料"
        ), "备份恢复失败"

    print("完整测试通过：提取、原文搜索、复制存储、重复检测、标签、删除、备份恢复均正常")


if __name__ == "__main__":
    main()
