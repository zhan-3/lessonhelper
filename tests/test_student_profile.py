import tempfile
import unittest
from pathlib import Path

from course_selection.student_profile import (
    create_student_profile,
    load_student_profile,
    save_student_profile,
    update_student_profile,
)


class StudentProfileTests(unittest.TestCase):
    def test_profile_round_trip_and_grade_normalization(self):
        profile = create_student_profile(grade="2025级")

        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "student-profile.json"
            save_student_profile(path, profile)

            self.assertEqual(load_student_profile(path), profile)
            self.assertEqual(profile.grade, "2025")

    def test_profile_can_be_enriched_without_guessing_unknown_facts(self):
        profile = create_student_profile(grade="2025")

        updated = update_student_profile(profile, major="计算机科学与技术")

        self.assertEqual(updated.grade, "2025")
        self.assertEqual(updated.major, "计算机科学与技术")
        self.assertEqual(updated.campus, "")

    def test_invalid_grade_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "四位入学年份"):
            create_student_profile(grade="大二")


if __name__ == "__main__":
    unittest.main()
