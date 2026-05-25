"""
Domain Models — مذكرتي Pro v17
Pure data classes with validation. No I/O, no side effects.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class SlideConfig:
    """Which slides to include"""
    cover: bool = True
    intro: bool = True
    plan: bool = True
    problem: bool = True
    objectives: bool = True
    importance: bool = True
    methodology: bool = True
    kpi: bool = True
    results: bool = True
    conclusion: bool = True
    recommendations: bool = True
    future: bool = True
    references: bool = True
    thankyou: bool = True

    @classmethod
    def from_dict(cls, d: dict) -> "SlideConfig":
        if not d:
            return cls()
        return cls(**{k: bool(v) for k, v in d.items() if k in cls.__dataclass_fields__})


@dataclass
class StatCard:
    label: str
    value: str
    unit: str = ""

    @classmethod
    def from_dict(cls, d: dict) -> Optional["StatCard"]:
        if not d or not d.get("label") or not d.get("value"):
            return None
        return cls(label=str(d["label"]), value=str(d["value"]), unit=str(d.get("unit", "")))


@dataclass
class Chapter:
    title: str
    pages: str = ""

    @classmethod
    def from_dict(cls, d: dict) -> Optional["Chapter"]:
        if not d or not d.get("title"):
            return None
        return cls(title=str(d["title"]), pages=str(d.get("pages", "")))


@dataclass
class PresentationRequest:
    """Validated, normalized input from frontend"""
    # Required
    student_name: str
    title_ar: str

    # Optional metadata
    title_en: str = ""
    supervisor: str = ""
    co_supervisor: str = ""
    institution: str = ""
    year: str = ""
    specialization: str = ""
    lang: str = "ar"
    engine: str = "canva"
    theme: str = "navy_gold"

    # Content
    intro_overview: str = ""
    intro_approach: str = ""
    main_problem: str = ""
    main_question: str = ""
    sub_questions: list[str] = field(default_factory=list)
    objectives: list[str] = field(default_factory=list)
    hypotheses: list[str] = field(default_factory=list)
    importance: list[str] = field(default_factory=list)
    reasons: str = ""
    methodology: str = ""
    sample_type: str = ""
    sample_size: str = ""
    tool: str = ""
    stats: list[StatCard] = field(default_factory=list)
    main_results: list[str] = field(default_factory=list)
    general_conclusion: str = ""
    recommendations: list[str] = field(default_factory=list)
    future_work: list[str] = field(default_factory=list)
    references: list[str] = field(default_factory=list)
    chapters: list[Chapter] = field(default_factory=list)

    # Slide toggles
    slides: SlideConfig = field(default_factory=SlideConfig)

    # يُحسب من pipeline قبل البناء
    _total_slides: int = 13

    VALID_THEMES = {
        'navy_gold', 'dark_teal', 'burgundy', 'forest',
        'midnight_purple', 'charcoal_orange', 'ice_blue',
        'sand_gold', 'slate_crimson', 'noir', 'atlas', 'sakura'
    }

    @classmethod
    def from_dict(cls, raw: dict) -> "PresentationRequest":
        def lst(k):
            return [str(x).strip() for x in (raw.get(k) or []) if str(x).strip()]

        theme = str(raw.get("theme", "navy_gold"))
        if theme not in cls.VALID_THEMES:
            theme = "navy_gold"

        stats = [s for s in (StatCard.from_dict(x) for x in (raw.get("stats") or [])) if s]
        chapters = [c for c in (Chapter.from_dict(x) for x in (raw.get("chapters") or [])) if c]
        slides = SlideConfig.from_dict(raw.get("slides") or {})

        return cls(
            student_name=str(raw.get("studentName", "")).strip(),
            title_ar=str(raw.get("titleAr", "")).strip(),
            title_en=str(raw.get("titleEn", "")).strip(),
            supervisor=str(raw.get("supervisor", "")).strip(),
            co_supervisor=str(raw.get("coSupervisor", "")).strip(),
            institution=str(raw.get("institution", "")).strip(),
            year=str(raw.get("year", "")).strip(),
            specialization=str(raw.get("specialization", "")).strip(),
            lang=str(raw.get("lang", "ar")),
            engine=str(raw.get("engine", "canva")),
            theme=theme,
            intro_overview=str(raw.get("introOverview", "")).strip(),
            intro_approach=str(raw.get("introApproach", "")).strip(),
            main_problem=str(raw.get("mainProblem", "")).strip(),
            main_question=str(raw.get("mainQuestion", "")).strip(),
            sub_questions=lst("subQuestions"),
            objectives=lst("objectives"),
            hypotheses=lst("hypotheses"),
            importance=lst("importance"),
            reasons=str(raw.get("reasons", "")).strip(),
            methodology=str(raw.get("methodology", "")).strip(),
            sample_type=str(raw.get("sampleType", "")).strip(),
            sample_size=str(raw.get("sampleSize", "")).strip(),
            tool=str(raw.get("tool", "")).strip(),
            stats=stats,
            main_results=lst("mainResults"),
            general_conclusion=str(raw.get("generalConclusion", "")).strip(),
            recommendations=lst("recommendations"),
            future_work=lst("futureWork"),
            references=lst("references"),
            chapters=chapters,
            slides=slides,
        )

    def validate(self) -> list[str]:
        errors = []
        if not self.student_name:
            errors.append("اسم الطالب مطلوب")
        if not self.title_ar:
            errors.append("عنوان المذكرة مطلوب")
        return errors
