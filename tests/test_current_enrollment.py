import unittest

from course_selection.current_enrollment import parse_current_enrollment_html

HTML = """
<table>
<tr><th>学年学期</th><th>课程代码</th><th>课程名称</th><th>课序号</th>
<th>课程类别</th><th>课程性质</th><th>开课院系</th><th>上课老师</th>
<th>上课地点</th><th>学分</th><th>总学时</th><th>教参数量</th><th>操作</th></tr>
<tr><td>2026-2027-1</td><td>ART1</td><td>舞蹈欣赏</td><td>01</td>
<td>文理通识-文化素质教育课</td><td>任选</td><td>人文学院</td><td>教师</td>
<td>教室</td><td>2.0</td><td>32</td><td>1</td><td>查看</td></tr>
<tr><td>2026-2027-1</td><td>MATH1</td><td>高等数学</td><td>01</td>
<td>专业基础课（包括大类平台课）</td><td>必修</td><td>理学院</td><td>教师</td>
<td>教室</td><td>4</td><td>64</td><td>1</td><td>查看</td></tr>
</table>
"""


class CurrentEnrollmentTests(unittest.TestCase):
    def test_parses_score_free_current_enrollment_facts(self):
        courses = parse_current_enrollment_html(HTML)

        self.assertEqual(len(courses), 2)
        self.assertEqual(courses[0].code, "ART1")
        self.assertEqual(courses[0].category, "文理通识-文化素质教育课")
        self.assertEqual(courses[0].nature, "任选")
        self.assertEqual(courses[0].credits, 2.0)

    def test_rejects_a_table_without_the_verified_headers(self):
        with self.assertRaisesRegex(ValueError, "已选课程表缺少"):
            parse_current_enrollment_html("<table><tr><th>课程名称</th></tr></table>")


if __name__ == "__main__":
    unittest.main()
