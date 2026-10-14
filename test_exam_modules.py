from __future__ import annotations

import hashlib
from dataclasses import replace
from docx import Document
from pathlib import Path
from tempfile import TemporaryDirectory

from exam_modules import (
    ExamService,
    analysis_markdown,
    answer_markdown,
    analyze_repeated_points,
    combine_exam_results,
    export_batch_analysis_markdown,
    export_batch_answer_markdown,
    export_batch_answer_word,
    export_answer_word,
    export_result_bundle,
    parse_exam,
    rank_answer_sources,
    export_individual_answers,
    answer_source_text,
)
from library_core import LibraryDatabase


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> None:
    with TemporaryDirectory() as temp:
        root = Path(temp)
        knowledge = root / "车站行车细则.md"
        knowledge.write_text(
            "# 调车作业\n\n"
            "转场调车时，为保证调车车列的制动力，遇到情况能及时停车，"
            "带车数在11辆及以上必须按《车站行车工作细则》规定连结一定数量的软管。本段为软件导出测试样例。\n\n"
            "摘结制动软管时，必须等车列停妥并得到调车长回示。\n",
            encoding="utf-8",
        )
        original_hash = digest(knowledge)
        exam = root / "测试试卷.md"
        exam.write_text(
            "# 测试试卷\n\n"
            "1. 转场调车时，为保证调车车列的制动力，带车数达到多少必须连结软管？\n"
            "A. 5辆 B. 8辆 C. 10辆 D. 11辆\n\n"
            "2. 转场调车时，为保证调车车列的制动力，遇到情况能及时停车，应执行什么规定？\n\n"
            "3. 月球背面的天气情况如何？\n",
            encoding="utf-8",
        )

        questions = parse_exam(exam)
        assert [item.number for item in questions] == ["1", "2", "3"]
        assert "A." not in questions[0].stem

        database = LibraryDatabase(root / "学习库.db", root / "资料文件")
        assert database.import_file(knowledge)[0] == "added"
        service = ExamService(database)
        hit = database.search("转场调车时")[0]
        service.add_marker(hit, "带车数在11辆及以上", "考点", "高频")
        assert len(service.list_markers()) == 1

        result = service.process_exam(exam)
        assert len(result.questions) == 3
        assert len(result.matches) == 2
        assert all(match.original_text in hit.content for match in result.matches.values())
        assert all(hit.content in match.source_text for match in result.matches.values())
        with database.connect() as connection:
            assert hit.content in connection.execute('SELECT source_text FROM exam_matches LIMIT 1').fetchone()[0]
        assert all("考点" in match.marker_types for match in result.matches.values())
        assert result.questions[2].id not in result.matches

        repeated, single = analyze_repeated_points(result, 0.78)
        assert repeated and len(repeated[0]["question_numbers"]) == 2
        assert not single

        answer = answer_markdown(result)
        assert answer.count("### 题目：") == 3
        assert hit.content.split("。")[0] in answer
        assert "本段为软件导出测试样例" not in answer
        assert "摘结制动软管时" not in answer
        assert "**【无】**" in answer
        assert "【来源：未检索到匹配文档】" in answer

        analysis = analysis_markdown(result, 0.78)
        assert "重复 2 次" in analysis and "第1题、第2题" in analysis
        assert "考点" in analysis
        word = export_answer_word(result, root / "答案.docx")
        assert word.exists() and word.stat().st_size > 0

        exam_two = root / "补考试卷.md"
        exam_two.write_text(
            "1. 转场调车时，为保证调车车列的制动力，带车数达到规定数量后应连结什么？\n",
            encoding="utf-8",
        )
        second_result = service.process_exam(exam_two)
        combined = combine_exam_results([result, second_result])
        repeated_batch, _single_batch = analyze_repeated_points(combined, 0.78)
        assert repeated_batch and len(repeated_batch[0]["question_numbers"]) == 3
        assert rank_answer_sources(combined)[0]['count'] == 3

        # Different selected sentences from one source chunk are still one cited source.
        first, second = result.questions[:2]
        altered = replace(result, matches=dict(result.matches))
        altered.matches[second.id] = replace(result.matches[second.id], original_text='同一来源片段的另一原句。')
        assert rank_answer_sources(altered)[0]['count'] == 2
        # Equal wording at a different source location must NOT inflate the exact ranking.
        altered.matches[second.id] = replace(result.matches[second.id], chunk_id=9999, source_location='另一个段落')
        assert [item['count'] for item in rank_answer_sources(altered)] == [1, 1]
        assert '最高引用次数：1 次' in analysis_markdown(altered, 0.5)
        assert '最高引用次数：1 次' in analysis_markdown(altered, 1.0)
        # Repeated display numbers count distinct question IDs, and equal paper names are disambiguated.
        renumbered = replace(result, questions=[replace(q, number='1') for q in result.questions])
        assert rank_answer_sources(renumbered)[0]['count'] == 2
        duplicate_names = combine_exam_results([result, replace(result)])
        assert rank_answer_sources(duplicate_names)[0]['count'] == 4
        assert '卷1' in duplicate_names.questions[0].number and '卷2' in duplicate_names.questions[3].number
        # Ranking uses distinct sources in descending frequency, with full original snapshots.
        mix = combine_exam_results([result, altered])
        assert [item['count'] for item in rank_answer_sources(mix)] == [3, 1]
        assert hit.content in rank_answer_sources(mix)[0]['text']
        empty = replace(result, matches={})
        assert rank_answer_sources(empty) == [] and '暂无考点数据' in analysis_markdown(empty)
        batch_md = export_batch_answer_markdown([result, second_result], root / "多卷答案.md")
        batch_word = export_batch_answer_word([result, second_result], root / "多卷答案.docx")
        batch_analysis = export_batch_analysis_markdown(
            [result, second_result], root / "多卷考点分析.md", 0.78
        )
        assert all(path.exists() and path.stat().st_size > 0 for path in (batch_md, batch_word, batch_analysis))
        assert "测试试卷·第1题" in batch_analysis.read_text(encoding="utf-8")
        bundle = export_result_bundle([result, second_result], root / "自动输出", 0.78, "TEST")
        assert len(bundle) == 7 and all(path.exists() for path in bundle)
        for index, paper in enumerate([result, second_result]):
            md_text = bundle[index * 3].read_text(encoding='utf-8')
            word_text = [p.text for p in Document(bundle[index * 3 + 1]).paragraphs]
            for question in paper.questions:
                heading = next(text for text in word_text if text.startswith('题目：') and question.stem in text)
                position = word_text.index(heading)
                assert word_text[position + 1] == '答案：'
                match = paper.matches.get(question.id)
                if match:
                    # Only the directly relevant original sentence is exported.
                    assert word_text[position + 2] == match.original_text
                    assert match.original_text in md_text
                    assert word_text[position + 3] == '出处：'
                    assert word_text[position + 4].startswith(f'【来源：{match.source_name} ·')
                else:
                    assert word_text[position + 2] == '【无】'
                    assert word_text[position + 4] == '【来源：未检索到匹配文档】'
            other = second_result if index == 0 else result
            assert not any(other.task_name in text for text in word_text)
        assert '月球背面' not in bundle[3].read_text(encoding='utf-8')
        assert answer_source_text(replace(next(iter(result.matches.values())), source_text=''))
        # Same names and repeated runs never overwrite each other's answers.
        duplicate_bundle = export_result_bundle([result, replace(result)], root / '自动输出', 0.78, 'TEST')
        assert duplicate_bundle[0].parent != bundle[0].parent
        assert len({path.name for path in duplicate_bundle}) == 7
        assert len(export_result_bundle([result], root / '自动输出')) == 3
        for fmt in ('md', 'docx'):
            separate = export_individual_answers([result, second_result], root / '手动输出', fmt)
            assert len(separate) == 2 and all(path.suffix == '.' + fmt and path.exists() for path in separate)
        assert digest(knowledge) == original_hash, "重点标记或试卷分析修改了原知识库文件"

        # Marker/task data must persist after creating a new service instance.
        assert len(ExamService(database).list_markers()) == 1
        # Existing pre-v2.4 task tables upgrade without dropping answers or source data.
        with database.connect() as connection:
            before = connection.execute('SELECT COUNT(*) FROM exam_matches').fetchone()[0]
            connection.execute('ALTER TABLE exam_matches DROP COLUMN source_text')
        ExamService(database)
        with database.connect() as connection:
            assert connection.execute('SELECT COUNT(*) FROM exam_matches').fetchone()[0] == before
            assert 'source_text' in {row['name'] for row in connection.execute('PRAGMA table_info(exam_matches)')}

    print("新增模块测试通过：试卷解析、原文检索、空答案、标记持久化、题目关联、重复考点和导出均正常")


if __name__ == "__main__":
    main()
