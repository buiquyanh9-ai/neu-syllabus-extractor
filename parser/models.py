"""
parser/models.py — Schema 2: Đề cương học phần

Changelog v3:
  - SyllabusRecord: thêm is_thuc_te + has_clo_table
  - _qa_check():
      * Không phạt "NO weeks" nếu is_thuc_te (Chuyên đề thực tế)
      * Không phạt "NO CLOs" nếu docx không có bảng CLO (has_clo_table=False)
"""
from dataclasses import dataclass, field
from typing import Optional
from datetime import datetime, timezone


# ── Sub-table dataclasses ─────────────────────────────────────────────────────

@dataclass
class Instructor:
    full_name: str
    email:     Optional[str] = None
    role_text: Optional[str] = None
    order_no:  int = 1

@dataclass
class Resource:
    resource_group: str
    citation_text:  str
    order_no:       int = 1

@dataclass
class Goal:
    goal_code:      str
    description_vi: str
    display_order:  int = 1

@dataclass
class CLO:
    clo_code:         str
    description_vi:   str
    goal_code:        Optional[str]   = None
    attainment_level: Optional[float] = None
    display_order:    int = 1

@dataclass
class Assessment:
    assessment_kind: str
    title:           str
    weight_percent:  Optional[float] = None
    timing_text:     Optional[str]   = None
    content_text:    Optional[str]   = None
    evaluation_tool: Optional[str]   = None
    clo_codes:       list = field(default_factory=list)
    display_order:   int  = 1

@dataclass
class Week:
    week_no:                      int
    session_order:                int = 1
    session_label:                Optional[str] = None
    content_text:                 Optional[str] = None
    reading_text:                 Optional[str] = None
    teaching_learning_activities: Optional[str] = None
    assessment_text:              Optional[str] = None
    clo_codes:                    list = field(default_factory=list)

@dataclass
class Requirement:
    requirement_group: str
    detail_text:       str
    order_no:          int = 1

@dataclass
class RubricCriterion:
    criterion_name: str
    clo_text:       Optional[str]   = None
    weight_percent: Optional[float] = None
    excellent_desc: Optional[str]   = None
    good_desc:      Optional[str]   = None
    average_desc:   Optional[str]   = None
    poor_desc:      Optional[str]   = None
    display_order:  int = 1

@dataclass
class Rubric:
    title:            str
    weight_percent:   Optional[float] = None
    overview_text:    Optional[str]   = None
    assessment_title: Optional[str]   = None
    display_order:    int = 1
    criteria:         list = field(default_factory=list)


# ── QA check ──────────────────────────────────────────────────────────────────

def _qa_check(record) -> dict:
    issues = []
    score  = 100

    if not record.course_code_snapshot:
        issues.append("MISSING course_code"); score -= 25
    if not record.course_name_vi_snapshot:
        issues.append("MISSING course_name"); score -= 15
    if record.credits is None:
        issues.append("MISSING credits"); score -= 10
    if not record.course_description:
        issues.append("MISSING course_description"); score -= 10

    # FIX: Không phạt NO CLOs nếu docx không có bảng CLO (format EN paragraph style)
    if not record.clos:
        if record.has_clo_table:
            issues.append("NO CLOs extracted — bảng CLO không parse được"); score -= 20
        # else: docx không có bảng CLO → bỏ qua (không phải lỗi parser)

    if not record.goals:
        issues.append("NO goals extracted — bảng Mục tiêu không parse được"); score -= 5
    if not record.instructors:
        issues.append("NO instructors — bảng giảng viên không parse được"); score -= 5
    if not record.resources:
        issues.append("NO resources — danh sách tài liệu không parse được"); score -= 5
    if not record.assessments:
        issues.append("NO assessments — bảng đánh giá không parse được"); score -= 5

    # FIX: Không phạt NO weeks nếu là Chuyên đề thực tế (field trip)
    if not record.weeks:
        if not record.is_thuc_te:
            issues.append("NO weeks — bảng kế hoạch dạy học không parse được"); score -= 5
        # else: Chuyên đề thực tế không có kế hoạch tuần → bình thường

    score = max(0, score)
    return {
        "is_ok":              len(issues) == 0,
        "issues":             issues,
        "completeness_score": score,
        "needs_review":       score < 60,
    }


# ── Main record ───────────────────────────────────────────────────────────────

@dataclass
class SyllabusRecord:
    # ── course_syllabi ────────────────────────────────────────────────────
    course_code_snapshot:    str
    course_name_vi_snapshot: str
    credits:                 Optional[int]
    course_description:      Optional[str]
    status:                  str = "draft"

    course_name_en_snapshot: Optional[str] = None
    decision_no:             Optional[str] = None
    level_name_vi:           Optional[str] = None
    managing_faculty_text:   Optional[str] = None
    faculty_address_text:    Optional[str] = None
    prerequisite_text:       Optional[str] = None
    class_hours:             Optional[int] = None
    self_study_hours:        Optional[int] = None

    # ── sub-tables ────────────────────────────────────────────────────────
    instructors:  list = field(default_factory=list)
    resources:    list = field(default_factory=list)
    goals:        list = field(default_factory=list)
    clos:         list = field(default_factory=list)
    assessments:  list = field(default_factory=list)
    weeks:        list = field(default_factory=list)
    requirements: list = field(default_factory=list)
    rubrics:      list = field(default_factory=list)

    # ── metadata ──────────────────────────────────────────────────────────
    source_file:  str = ""
    extracted_at: str = field(
        default_factory=lambda: datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    )

    # FIX: flags để QA thông minh hơn
    is_thuc_te:    bool = False   # True = Chuyên đề thực tế, không cần bảng tuần
    has_clo_table: bool = True    # False = docx không có bảng CLO (không phải lỗi parser)

    def to_dict(self) -> dict:
        qa = _qa_check(self)
        return {
            "syllabus": {
                "_course_id":              None,
                "_source_document_id":     None,
                "course_code_snapshot":    self.course_code_snapshot,
                "course_name_vi_snapshot": self.course_name_vi_snapshot,
                "course_name_en_snapshot": self.course_name_en_snapshot,
                "credits":                 self.credits,
                "class_hours":             self.class_hours,
                "self_study_hours":        self.self_study_hours,
                "level_name_vi":           self.level_name_vi,
                "decision_no":             self.decision_no,
                "managing_faculty_text":   self.managing_faculty_text,
                "faculty_address_text":    self.faculty_address_text,
                "prerequisite_text":       self.prerequisite_text,
                "course_description":       self.course_description,
                "status":                  self.status,
            },
            "instructors": [
                {
                    "full_name": i.full_name,
                    "email":     i.email,
                    "role_text": i.role_text,
                    "order_no":  i.order_no,
                }
                for i in self.instructors
            ],
            "resources": [
                {
                    "resource_group": r.resource_group,
                    "citation_text":  r.citation_text,
                    "order_no":       r.order_no,
                }
                for r in self.resources
            ],
            "goals": [
                {
                    "goal_code":      g.goal_code,
                    "description_vi": g.description_vi,
                    "description_en": None,
                    "display_order":  g.display_order,
                }
                for g in self.goals
            ],
            "clos": [
                {
                    "clo_code":         c.clo_code,
                    "goal_code":        c.goal_code,
                    "description_vi":   c.description_vi,
                    "description_en":   None,
                    "attainment_level": c.attainment_level,
                    "display_order":    c.display_order,
                }
                for c in self.clos
            ],
            "assessments": [
                {
                    "assessment_kind": a.assessment_kind,
                    "title":           a.title,
                    "weight_percent":  a.weight_percent,
                    "timing_text":     a.timing_text,
                    "content_text":    a.content_text,
                    "evaluation_tool": a.evaluation_tool,
                    "clo_codes":       a.clo_codes,
                    "display_order":   a.display_order,
                }
                for a in self.assessments
            ],
            "weeks": [
                {
                    "week_no":                      w.week_no,
                    "session_order":                w.session_order,
                    "session_label":                w.session_label,
                    "content_text":                 w.content_text,
                    "reading_text":                 w.reading_text,
                    "teaching_learning_activities": w.teaching_learning_activities,
                    "assessment_text":              w.assessment_text,
                    "clo_codes":                    w.clo_codes,
                }
                for w in self.weeks
            ],
            "requirements": [
                {
                    "requirement_group": r.requirement_group,
                    "detail_text":       r.detail_text,
                    "order_no":          r.order_no,
                }
                for r in self.requirements
            ],
            "rubrics": [
                {
                    "title":            rb.title,
                    "weight_percent":   rb.weight_percent,
                    "overview_text":    rb.overview_text,
                    "assessment_title": rb.assessment_title,
                    "display_order":    rb.display_order,
                    "criteria": [
                        {
                            "criterion_name": cr.criterion_name,
                            "clo_text":       cr.clo_text,
                            "weight_percent": cr.weight_percent,
                            "excellent_desc": cr.excellent_desc,
                            "good_desc":      cr.good_desc,
                            "average_desc":   cr.average_desc,
                            "poor_desc":      cr.poor_desc,
                            "display_order":  cr.display_order,
                        }
                        for cr in rb.criteria
                    ],
                }
                for rb in self.rubrics
            ],
            "_meta": {
                "source_file":    self.source_file,
                "extracted_at":   self.extracted_at,
                "is_thuc_te":     self.is_thuc_te,
                "has_clo_table":  self.has_clo_table,
            },
            "_qa": qa,
        }
