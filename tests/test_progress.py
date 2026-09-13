import unittest
from pathlib import Path

from course_progress.baselines import requirement_baseline
from course_progress.progress import (
    AcademicRecord,
    CompletedCourse,
    Requirement,
    RequirementBaseline,
    apply_course_label_estimates,
    apply_projected_course_estimates,
    apply_recognized_credit_estimates,
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

    def test_user_course_labels_change_only_estimated_subconstraints(self):
        baseline = baseline_from_definition(
            requirement_baseline("basic-graduation-reference-v1")
        )
        records = (
            AcademicRecord("2024秋季", "C01", "文化课程", "任选", "文理通识-文化素质教育课", 2.0, True),
            AcademicRecord("2025春季", "C01", "文化课程", "任选", "文理通识-文化素质教育课", 2.0, True),
            AcademicRecord("2025春季", "O01", "外专业课程", "任选", "跨专业发展课程", 10.0, True),
        )
        report = evaluate_progress(records, baseline)
        confirmed = confirmed_progress_items(report, data_complete=True)

        estimated = apply_course_label_estimates(
            confirmed, report.unclassified_courses,
            (
                {"course_identity": "C01", "d_category": True, "four_histories": True, "outside_track": ""},
                {"course_identity": "O01", "d_category": False, "four_histories": False, "outside_track": "track-a"},
            ),
            "track-a",
        )
        items = {item["key"]: item for item in estimated}

        self.assertEqual(2, items["cultural_quality"]["confirmed_amount"])
        self.assertEqual(2, items["cultural_quality"]["estimated_amount"])
        self.assertEqual(2, items["cultural_quality_d"]["estimated_amount"])
        self.assertEqual(1, items["four_histories"]["estimated_amount"])
        self.assertEqual(10, items["outside_major_elective"]["estimated_amount"])

        unknown_track = apply_course_label_estimates(
            confirmed, report.unclassified_courses,
            ({"course_identity": "O01", "d_category": True, "four_histories": False, "outside_track": ""},),
            "track-a",
        )
        outside = next(item for item in unknown_track if item["key"] == "outside_major_elective")
        self.assertEqual(["O01"], outside["unknown_track_course_identities"])

    def test_projected_sources_use_confirmed_enrolled_queue_priority_and_combination_dedupe(self):
        baseline_definition = requirement_baseline("basic-graduation-reference-v1")
        baseline = baseline_from_definition(baseline_definition)
        report = evaluate_progress((
            AcademicRecord("2025春季", "C1", "已修创新", "任选", "创新研修课", 2.0, True),
        ), baseline)
        confirmed = confirmed_progress_items(report, data_complete=True)
        declared = apply_recognized_credit_estimates(confirmed, baseline_definition, ({
            "identity": "declaration-1", "category": "innovation", "credits": 1,
            "linked_course_identity": "", "note": "合成申报",
        },))
        labeled = apply_course_label_estimates(declared, (), (), "")

        projected = apply_projected_course_estimates(
            labeled,
            baseline_definition,
            enrolled_courses=(
                {"code": "C1", "name": "已修创新", "category": "创新研修课", "credits": 2},
                {"code": "E1", "name": "本学期创新", "category": "创新研修课", "credits": 1},
            ),
            queued_courses=(
                {"course_code": "E1", "course_name": "重复队列创新", "category": "创新研修课", "credits": 1},
                {"course_code": "Q1", "course_name": "队列实践", "category": "社会实践", "credits": 1},
                {"course_code": "Q2", "course_name": "待核验文化", "category": "文理通识-文化素质教育课", "credits": 2},
                {"course_code": "Q3", "course_name": "明确体系课程", "category": "跨专业发展课程", "credits": 2},
                {"course_code": "Q4", "course_name": "待核验文化", "category": "文理通识-文化素质教育课", "credits": 2},
                {"course_code": "Q5", "course_name": "体系未知课程", "category": "跨专业发展课程", "credits": 2},
                {"course_code": "BAD", "course_name": "异常学分", "category": "创新研修课", "credits": float("nan")},
            ),
            labels=(
                {"course_identity": "Q2", "d_category": True, "four_histories": True, "outside_track": ""},
                {"course_identity": "Q3", "d_category": False, "four_histories": False, "outside_track": "track-a"},
            ), selected_track="track-a",
        )
        items = {item["key"]: item for item in projected}

        self.assertEqual(2, items["innovation"]["confirmed_amount"])
        self.assertEqual(1, items["innovation"]["enrolled_amount"])
        self.assertEqual(0, items["innovation"]["queued_amount"])
        self.assertIn("BAD", {
            pending["identity"] for pending in items["innovation"]["pending_verification"]
            if pending["kind"] == "invalid_credits"
        })
        self.assertEqual(1, items["innovation"]["declared_amount"])
        self.assertEqual(4, items["innovation"]["estimated_amount"])
        self.assertEqual(1, items["social_practice"]["queued_amount"])
        self.assertEqual(5, items["innovation_and_practice"]["estimated_amount"])
        self.assertEqual("unknown", items["innovation_and_practice"]["estimated_condition_status"])
        self.assertEqual(4, items["cultural_quality"]["queued_amount"])
        self.assertEqual(2, items["cultural_quality_d"]["queued_amount"])
        self.assertEqual(1, items["four_histories"]["queued_amount"])
        self.assertEqual("Q4", items["cultural_quality_d"]["pending_verification"][0]["identity"])
        self.assertEqual("Q4", items["four_histories"]["pending_verification"][0]["identity"])
        self.assertEqual(2, items["outside_major_elective"]["queued_amount"])
        self.assertIn("Q5", items["outside_major_elective"]["unknown_track_course_identities"])

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
