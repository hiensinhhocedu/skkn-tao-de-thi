
export type ExamType = 'Thường xuyên' | 'Giữa Kì' | 'Cuối Kì';
export type Subject = 'Toán' | 'Vật Lí' | 'Hóa học' | 'Sinh học' | 'Lịch sử' | 'Địa Lí' | 'Kinh tế-Pháp luật' | 'Công Nghệ Nông Nghiệp' | 'Công Nghệ Công Nghiệp' | 'Hoạt động trải nghiệm' | 'Thể dục – Quốc phòng';
export type Grade = '10' | '11' | '12';
export type GenerationMethod = 'topic' | 'objective';
export type ExamStructure = 'Trắc nghiệm' | 'Tự Luận' | 'Trắc nghiệm + Tự Luận';

export interface Objective {
  id: number;
  topic: string;
  requirement: string;
}

export interface QuestionCounts {
  objective: number;
  trueFalse: number;
  shortAnswer: number;
  essay: number;
}

export interface CognitiveLevels {
  knowledge: number;
  comprehension: number;
  application: number;
}

export interface TopicRow {
    id: number;
    topic: string;
    mcq_know: number;
    mcq_comp: number;
    mcq_app: number;
    tf_know: number;
    tf_comp: number;
    tf_app: number;
    sa_know: number;
    sa_comp: number;
    sa_app: number;
    essay_know: number;
    essay_comp: number;
    essay_app: number;
    total_know: number;
    total_comp: number;
    total_app: number;
    total_sum: number;
}

export interface SummaryRow {
    label: string;
    mcq_know: number;
    mcq_comp: number;
    mcq_app: number;
    tf_know: number;
    tf_comp: number;
    tf_app: number;
    sa_know: number;
    sa_comp: number;
    sa_app: number;
    essay_know: number;
    essay_comp: number;
    essay_app: number;
    total_know: number;
    total_comp: number;
    total_app: number;
    total_sum: number;
}

export interface MatrixData {
    topicRows: TopicRow[];
    summaryRows: SummaryRow[];
}

export interface QuestionNumbersByLevel {
    knowledge: string;
    comprehension: string;
    application: string;
}

export interface SpecificationTopic {
    id: number;
    content: string;
    questionNumbers: {
        mcq: QuestionNumbersByLevel;
        tf: QuestionNumbersByLevel;
        sa: QuestionNumbersByLevel;
        essay: QuestionNumbersByLevel;
    };
    requirements: {
        knowledge: string;
        comprehension: string;
        application: string;
    };
    mcq_know: number;
    mcq_comp: number;
    mcq_app: number;
    tf_know: number;
    tf_comp: number;
    tf_app: number;
    sa_know: number;
    sa_comp: number;
    sa_app: number;
    essay_know: number;
    essay_comp: number;
    essay_app: number;
}

export interface SpecificationSummaryRow {
    label: string;
    mcq_know: number;
    mcq_comp: number;
    mcq_app: number;
    tf_know: number;
    tf_comp: number;
    tf_app: number;
    sa_know: number;
    sa_comp: number;
    sa_app: number;
    essay_know: number;
    essay_comp: number;
    essay_app: number;
}


export interface SpecificationData {
    topics: SpecificationTopic[];
    summaryRows: SpecificationSummaryRow[];
}

export interface ExamSubQuestion {
    text: string;
    level: string;
}

export interface ExamQuestion {
    type: string;
    level: string;
    question: string;
    options?: string[]; // For 'Trắc nghiệm khách quan'
    subQuestions?: ExamSubQuestion[]; // For 'Đúng/Sai'
    answer: string;
    explanation?: string;
}

export interface ExamData {
    header: {
        examType: string;
        subject: string;
        time: string;
    };
    questions: ExamQuestion[];
}

export interface MatricesData {
    matrix: MatrixData;
    specification: SpecificationData;
}
