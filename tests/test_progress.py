import unittest
from pathlib import Path

from course_progress.baselines import requirement_baseline
from course_progress.progress import (
    AcademicRecord,
    CompletedCourse,
    Requirement,
    RequirementBaseline,
    assess_progress,
    baseline_from_definition,
    calculate_progress,
    confirmed_progress_items,
    evaluate_progress,
    parse_grade_html,
    parse_requirements,
)


class ProgressTests(unittest.TestCase):
    def test_grade_html_exposes_completion_without_retaining_scores(self):
        records = parse_grade_html(
            """
            <table><tr><th>学年学期</th><th>课程代码</th><th>课程名称</th>
            <th>课程性质</th><th>课程类别</th><th>学分</th><th>最终成绩</th></tr>
            <tr><td>2025秋季</td><td>A01</td><td>四史专题</td><td>任选</td>
            <td>文理通识-文化素质教育课</td><td>2.0</td><td>85</td></tr>
            <tr><td>2025秋季</td><td>A02</td><td>未通过课程</td><td>任选</td>
            <td>外专业选修</td><td>1.0</td><td>55</td></tr></table>
            """
        )
        self.assertEqual(
            records,
            (
                AcademicRecord("2025秋季", "A01", "四史专题", "任选", "文理通识-文化素质教育课", 2.0, True),
                AcademicRecord("2025秋季", "A02", "未通过课程", "任选", "外专业选修", 1.0, False),
            ),
        )

    def test_evaluates_only_unique_passed_non_mandatory_courses(self):
        baseline = RequirementBaseline(
            version="guide-2026",
            requirements=(Requirement("cultural_quality", "文化素质课程", 8.0),),
            category_mapping={"文理通识-文化素质教育课": "cultural_quality"},
        )
        records = (
            AcademicRecord("2025秋季", "A01", "四史专题", "任选", "文理通识-文化素质教育课", 2.0, True),
            AcademicRecord("2026春季", "A01", "四史专题", "任选", "文理通识-文化素质教育课", 2.0, True),
            AcademicRecord("2025秋季", "A02", "未通过课程", "任选", "文理通识-文化素质教育课", 3.0, False),
            AcademicRecord("2025秋季", "A03", "必修课", "必修", "文理通识-文化素质教育课", 3.0, True),
        )

        report = evaluate_progress(records, baseline)

        self.assertEqual(report.baseline_version, "guide-2026")
        self.assertEqual(report.progress[0].completed_credits, 2.0)
        self.assertEqual(report.progress[0].remaining_credits, 6.0)
        self.assertEqual([course.code for course in report.progress[0].courses], ["A01"])

    def test_conflicting_course_identity_is_not_counted(self):
        baseline = RequirementBaseline(
            version="guide-2026",
            requirements=(Requirement("cultural_quality", "文化素质课程", 8.0),),
            category_mapping={"文理通识-文化素质教育课": "cultural_quality"},
        )
        records = (
            AcademicRecord("2025秋季", "A01", "四史专题", "任选", "文理通识-文化素质教育课", 2.0, True),
            AcademicRecord("2026春季", "A01", "四史专题", "任选", "文理通识-文化素质教育课", 1.0, True),
        )

        report = evaluate_progress(records, baseline)

        self.assertEqual(report.progress[0].completed_credits, 0.0)
        self.assertEqual([conflict.identity for conflict in report.conflicts], ["A01"])

    def test_unknown_category_remains_unclassified(self):
        baseline = RequirementBaseline(
            version="guide-2026",
            requirements=(Requirement("cultural_quality", "文化素质课程", 8.0),),
            category_mapping={},
        )
        record = AcademicRecord(
            "2025秋季", "A01", "未知选修课", "任选", "新课程类别", 2.0, True
        )

        report = evaluate_progress((record,), baseline)

        self.assertEqual(report.progress[0].completed_credits, 0.0)
        self.assertEqual([course.name for course in report.unclassified_courses], ["未知选修课"])

    def test_combined_requirement_sums_its_child_categories(self):
        baseline = RequirementBaseline(
            version="guide-2026",
            requirements=(
                Requirement(
                    "innovation_and_practice",
                    "创新创业 + 社会实践",
                    6.0,
                    contribution_keys=("innovation", "social_practice"),
                ),
                Requirement("innovation", "创新创业", 4.0),
                Requirement("social_practice", "社会实践", 1.0),
            ),
            category_mapping={"创新研修课": "innovation", "社会实践": "social_practice"},
        )
        records = (
            AcademicRecord("2025秋季", "I01", "创新课程", "任选", "创新研修课", 4.0, True),
            AcademicRecord("2026春季", "S01", "社会实践", "任选", "社会实践", 1.0, True),
        )

        report = evaluate_progress(records, baseline)
        progress = {item.requirement.key: item for item in report.progress}

        self.assertEqual(progress["innovation_and_practice"].completed_credits, 5.0)
        self.assertEqual(progress["innovation_and_practice"].remaining_credits, 1.0)
        self.assertEqual(progress["innovation"].remaining_credits, 0.0)
        self.assertEqual(progress["social_practice"].remaining_credits, 0.0)

    def test_basic_reference_baseline_assesses_confirmed_minimums_and_combination(self):
        baseline = baseline_from_definition(
            requirement_baseline("basic-graduation-reference-v1")
        )
        records = (
            AcademicRecord("2025秋季", "M01", "专业选修", "任选", "本专业选修", 2.0, True),
            AcademicRecord("2025秋季", "I01", "创新课程", "任选", "创新研修课", 4.0, True),
            AcademicRecord("2025秋季", "S01", "社会实践", "任选", "社会实践", 1.0, True),
        )

        report = evaluate_progress(records, baseline)
        assessed = {item.progress.requirement.key: item for item in assess_progress(report, data_complete=True)}

        self.assertEqual("not_satisfied", assessed["major_elective"].state)
        self.assertEqual(2.0, assessed["major_elective"].confirmed_amount)
        self.assertEqual(1.0, assessed["major_elective"].confirmed_gap)
        self.assertEqual("satisfied", assessed["innovation"].state)
        self.assertEqual("satisfied", assessed["social_practice"].state)
        self.assertEqual(5.0, assessed["innovation_and_practice"].confirmed_amount)
        self.assertEqual("unknown", assessed["innovation_and_practice"].state)
        self.assertEqual("recognized_credit_coverage_missing", assessed["innovation_and_practice"].reason)
        payload = confirmed_progress_items(report, data_complete=True)
        self.assertEqual(8, len(payload))
        innovation = next(item for item in payload if item["key"] == "innovation")
        self.assertEqual(4.0, innovation["minimum"])
        self.assertEqual("manual-supplement", innovation["source"])
        self.assertEqual("satisfied", innovation["condition_status"])
        self.assertTrue(innovation["courses"])

    def test_basic_reference_baseline_keeps_unproven_subconstraints_unknown(self):
        baseline = baseline_from_definition(
            requirement_baseline("basic-graduation-reference-v1")
        )
        records = (
            AcademicRecord(
                "2025秋季", "C01", "四史专题", "任选",
                "文理通识-文化素质教育课", 8.0, True,
            ),
            AcademicRecord(
                "2025秋季", "O01", "跨专业课程", "任选",
                "跨专业发展课程", 10.0, True,
            ),
        )

        report = evaluate_progress(records, baseline)
        assessed = {item.progress.requirement.key: item for item in assess_progress(report, data_complete=True)}

        self.assertEqual(8.0, assessed["cultural_quality"].confirmed_amount)
        self.assertEqual("unknown", assessed["cultural_quality"].state)
        self.assertEqual("required_subconstraint_unknown", assessed["cultural_quality"].reason)
        self.assertEqual(0.0, assessed["cultural_quality_d"].confirmed_amount)
        self.assertEqual("unknown", assessed["cultural_quality_d"].state)
        self.assertEqual(0.0, assessed["four_histories"].confirmed_amount)
        self.assertEqual("unknown", assessed["four_histories"].state)
        self.assertEqual("unknown", assessed["outside_major_elective"].state)
        self.assertEqual("outside_major_track_unconfirmed", assessed["outside_major_elective"].reason)

    def test_cultural_subconstraint_courses_contribute_once_to_the_parent_total(self):
        definition = requirement_baseline("basic-graduation-reference-v1")
        definition["category_mapping"].update({
            "明确D类四史": ["cultural_quality_d", "four_histories"],
        })
        baseline = baseline_from_definition(definition)
        records = (
            AcademicRecord("2025秋季", "DH01", "党史专题", "任选", "明确D类四史", 2.0, True),
        )

        report = evaluate_progress(records, baseline)
        progress = {item.requirement.key: item for item in report.progress}

        self.assertEqual(2.0, progress["cultural_quality"].completed_amount)
        self.assertEqual(2.0, progress["cultural_quality_d"].completed_amount)
        self.assertEqual(1.0, progress["four_histories"].completed_amount)
        self.assertEqual(1, len(progress["cultural_quality"].courses))

    def test_incomplete_grade_data_cannot_prove_a_deficit(self):
        baseline = baseline_from_definition(
            requirement_baseline("basic-graduation-reference-v1")
        )
        report = evaluate_progress((), baseline)
        assessed = {item.progress.requirement.key: item for item in assess_progress(report, data_complete=False)}

        self.assertEqual("unknown", assessed["major_elective"].state)
        self.assertEqual("grade_data_incomplete", assessed["major_elective"].reason)

    def test_parses_requirements_from_extracted_guide(self):
        requirements = parse_requirements(
            Path("docs/校园培养方案解读（2026年版）.md")
        )
        values = {item.key: item.minimum_credits for item in requirements}
        self.assertEqual(values["major_elective"], 3.0)
        self.assertEqual(values["outside_major_elective"], 10.0)
        self.assertEqual(values["cultural_quality"], 8.0)
        self.assertNotIn("innovation", values)
        self.assertEqual(values["innovation_and_practice"], 6.0)
        self.assertEqual(values["social_practice"], 1.0)

    def test_excludes_mandatory_courses_and_keeps_course_details(self):
        requirements = parse_requirements(
            Path("docs/校园培养方案解读（2026年版）.md")
        )
        courses = [
            CompletedCourse("A", "必修课", "必修", "文理通识-文化素质教育课", 3.0),
            CompletedCourse("B", "四史专题", "任选", "文理通识-文化素质教育课", 2.0),
            CompletedCourse("C", "创新课", "任选", "创新研修课", 1.5),
        ]
        progress = {item.requirement.key: item for item in calculate_progress(requirements, courses)}
        self.assertEqual(progress["cultural_quality"].completed_credits, 2.0)
        self.assertEqual(progress["cultural_quality"].remaining_credits, 6.0)
        self.assertEqual(progress["cultural_quality"].courses[0].name, "四史专题")
        self.assertEqual(progress["innovation_and_practice"].completed_credits, 1.5)


if __name__ == "__main__":
    unittest.main()
