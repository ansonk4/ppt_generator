# PPT 報告生成器

**語言:** [English](README.md) | [中文](README_zh.md)

此項目根據香港中學文憑試（DSE）學生關於未來升學及就業計劃的調查數據，自動生成 PowerPoint 簡報。系統會處理數據並以可視化形式呈現於簡報中。

## 使用方法

您可以直接在線上使用 PPT 報告生成器：[https://pptgenerator-uqzlepmjiojt6h6mowtmk7.streamlit.app/](https://pptgenerator-uqzlepmjiojt6h6mowtmk7.streamlit.app/)，並跳過以下步驟。

若想在本地運行，請按以下步驟操作：

1. 確保使用 Python 3.10 或以上版本

2. 複製此代碼庫：

   ```bash
   git clone https://github.com/ansonk4/ppt_generator.git
   cd ppt
   ```

3. 安裝依賴項：

   ```bash
   pip install -r requirements.txt
   ```

4. 運行 Streamlit 網頁界面：
   ```bash
   streamlit run src/streamlit.py
   ```

## 輸入數據格式

系統需要一個包含特定欄位及數值格式的 Excel 文件。[樣本數據文件](sample_data/sample.xlsx)可供參考，建議按此格式準備您的數據。

### 必需欄位

以下欄位必須包含在您的 Excel 文件中：

#### 人口統計與背景

- `性別` (Gender)
- `Banding` (學校組別)
- `學校編號` (School ID)
- `父母教育程度` (Parental Education Level)
- `高中選修學科` (High School Elective Subjects)
- `中文成績` (Chinese Score)
- `英文成績` (English Score)
- `數學成績` (Math Score)

#### DSE 後計劃

- `大學` (University)
- `副學士` (Associate Degree)
- `文憑` (Diploma)
- `高級文憑` (Higher Diploma)
- `工作` (Work)
- `工作假期` (Working Holiday)
- `其他` (Other)
- `試後計劃` (Post-Exam Plans)

#### 升學地點

- `香港` (Hong Kong)
- `內地` (Mainland China)
- `亞洲` (Asia)
- `歐美澳` (Europe/America/Australia)

#### 心儀大學（二元指標）

- `浸會大學` (HKBU)
- `中文大學` (CUHK)
- `城市大學` (CityU)
- `教育大學` (EdUHK)
- `恒生大學` (HSUHK)
- `香港大學` (HKU)
- `嶺南大學` (Lingnan)
- `都會大學` (MUHK)
- `理工大學` (PolyU)
- `聖方濟各大學` (SFJU)
- `樹仁大學` (HKSYU)
- `科技大學` (UST)
- `自資學院` (Self-financed Colleges)

#### 學科偏好

- `希望修讀` (Wish to Study)
- `希望修讀_A` (Wish to Study A)
- `希望修讀_B` (Wish to Study B)
- `不希望修讀` (Do Not Wish to Study)
- `不希望修讀_A` (Do Not Wish to Study A)
- `不希望修讀_B` (Do Not Wish to Study B)

#### 職業輔導活動

- `大學入學講座_A` (University Admission Seminars A)
- `升學展覽_A` (Education Fairs A)
- `職業博覽_A` (Career Fairs A)
- `生涯規劃_A` (Career Planning A)
- `團體師友_A` (Group Mentoring A)
- `工作影子_A` (Job Shadowing A)
- `大學入學講座_B` (University Admission Seminars B)
- `升學展覽_B` (Education Fairs B)
- `職業博覽_B` (Career Fairs B)
- `生涯規劃_B` (Career Planning B)
- `團體師友_B` (Group Mentoring B)
- `工作影子_B` (Job Shadowing B)

#### 決策因素

- `學科知識` (Subject Knowledge)
- `院校因素` (Institutional Factors)
- `大學學費` (University Tuition Fees)
- `助學金` (Financial Aid)
- `主要行業` (Main Industries)
- `朋輩老師` (Peers and Teachers)
- `家庭因素` (Family Factors)
- `預期收入` (Expected Income)
- `DSE成績` (DSE Results)
- `高中選修科目` (High School Elective Subjects)

#### 大灣區認知

- `大灣區了解` (GBA Understanding)
- `公社科` (Civic and Social Sciences)
- `內地考察` (Mainland Visits)
- `政府資訊` (Government Information)
- `新聞媒體` (News Media)
- `網上資訊` (Online Information)
- `內地交流` (Mainland Exchange)
- `校內講座` (School Lectures)
- `朋輩及老師` (Peers and Teachers)

#### 大灣區就業考慮

- `個人興趣及性格_gba` (Personal Interests and Personality)
- `個人能力_gba` (Personal Abilities)
- `晉升機會_gba` (Promotion Opportunities)
- `工作性質_gba` (Nature of Work)
- `行業前景_gba` (Industry Prospects)
- `工作環境_gba` (Work Environment)
- `工作量_gba` (Workload)
- `薪水福利_gba` (Salary and Benefits)
- `生活成本_gba` (Cost of Living)
- `國家貢獻_gba` (National Contribution)

#### 工作偏好

- `工作地方` (Work Location)
- `個人能力_B` (Personal Abilities)
- `個人興趣性格_B` (Personal Interests and Personality)
- `成就感_B` (Sense of Achievement)
- `家庭因素_B` (Family Factors)
- `人際關係_B` (Interpersonal Relationships)
- `工作性質_B` (Nature of Work)
- `工作模式_B` (Work Mode)
- `工作量_B` (Workload)
- `工作環境_B` (Work Environment)
- `薪水及褔利_B` (Salary and Benefits)
- `晉升機會_B` (Promotion Opportunities)
- `發展前景_B` (Development Prospects)
- `社會貢獻_B` (Social Contribution)
- `社會地位_B` (Social Status)
- `希望從事` (Hope to Engage In)
- `希望從事_A` (Hope to Engage In A)
- `希望從事_B` (Hope to Engage In B)
- `不希望從事` (Do Not Hope to Engage In)
- `不希望從事_A` (Do Not Hope to Engage In A)
- `不希望從事_B` (Do Not Hope to Engage In B)
- `從事相關工作` (Engage in Related Work)

#### STEM 教育

- `參加STEM` (Participate in STEM)
- `STEM影響職業選擇程度` (STEM Influence on Career Choice)
- `領導能力` (Leadership)
- `團隊合作` (Teamwork)
- `創新思維` (Innovative Thinking)
- `科學知識` (Scientific Knowledge)
- `解難能力` (Problem Solving)

---

### 數值驗證規則

以下欄位必須僅包含指定列表中的數值：

#### `性別` (Gender)

- `男` (Male)
- `女` (Female)

#### `Banding` (學校組別)

- `Band 1`
- `Band 2`
- `Band 3`

#### 職業輔導活動

欄位：`大學入學講座_A`, `升學展覽_A`, `職業博覽_A`, `生涯規劃_A`, `團體師友_A`, `工作影子_A`

- `有` (Yes)
- `沒有` (No)

#### `大灣區了解` (GBA Understanding)

- `完全不了解` (Completely Unfamiliar)
- `不太了解` (Not Very Familiar)
- `了解` (Familiar)
- `非常了解` (Very Familiar)

#### 大灣區參與活動

欄位：`公社科`, `內地考察`, `政府資訊`, `新聞媒體`, `網上資訊`, `內地交流`, `校內講座`, `朋輩及老師`

- `曾經 / 希望參與` (Have/Wish to Participate)
- `沒有 / 不會參與` (Have Not/Won't Participate)

#### `工作地方` (Work Location)

- `香港` (Hong Kong)
- `內地` (Mainland China)
- `國外 - 亞洲` (Abroad - Asia)
- `國外 - 歐美澳` (Abroad - Europe/America/Australia)

#### 工作重要性評分

欄位：`個人能力_B`, `個人興趣性格_B`, `成就感_B`, `家庭因素_B`, `人際關係_B`,
`工作性質_B`, `工作模式_B`, `工作量_B`, `工作環境_B`, `薪水及褔利_B`,
`晉升機會_B`, `發展前景_B`, `社會貢獻_B`, `社會地位_B`

- `十分重要` (Very Important)
- `重要` (Important)
- `不太重要` (Not Very Important)
- `不重要` (Not Important)

#### `參加STEM` (Participate in STEM)

- `有` (Yes)
- `沒有` (No)

#### 學科成績

欄位：`中文成績`, `英文成績`, `數學成績`

- `< 25 分` (Less than 25 points)
- `25-49 分` (25-49 points)
- `50-75 分` (50-75 points)
- `> 75 分` (More than 75 points)

#### `從事相關工作` (Engage in Related Work)

- `絕對不會` (Absolutely Not)
- `可能不會` (Probably Not)
- `不確定` (Uncertain)
- `可能會` (Probably Will)
- `絕對會` (Absolutely Will)

#### 心儀大學

欄位：`浸會大學`, `中文大學`, `城市大學`, `教育大學`, `恒生大學`, `香港大學`,
`嶺南大學`, `都會大學`, `理工大學`, `聖方濟各大學`, `樹仁大學`, `科技大學`, `自資學院`

- `1` (Yes/Selected)
- `0` (No/Not Selected)

### 大灣區

欄位：`個人興趣及性格_gba`, `個人能力_gba`, `晉升機會_gba`, `工作性質_gba`, `行業前景_gba`, `工作環境_gba`, `工作量_gba`, `薪水福利_gba`, `生活成本_gba`, `國家貢獻_gba`

- `1` (Yes/Selected)
- `0` (No/Not Selected)

---

**以下欄位必須至少包含（但不限於）指定列表中的一項數值：**

欄位：`希望修讀`, `希望修讀_A` , `希望修讀_B` , `不希望修讀`, `不希望修讀_A`, `不希望修讀_B`

- `電腦工程`
- `電腦科學`
- `數學`
- `金融`
- `法律`

欄位：`希望從事`, `希望從事_A` ,`希望從事_B` ,`不希望從事`,`不希望從事_A`,`不希望從事_B`:

- `資訊科技`,
- `電腦工程`,
- `銀行/金融`,
- `創業`

### 特殊數值

所有欄位均接受 `999` 作為特殊數值，表示「不適用」或「跳過」。此數值在數據處理時會轉換為 NaN。

## 未指定欄位

- 未列於[必需欄位](#必需欄位)的欄位不會被系統使用或驗證，可刪除或按需修改。
- 未包含在[數值驗證規則](#數值驗證規則)的欄位可包含任何數值，這些數值將直接顯示於生成的 PowerPoint 中。
  建議遵循[樣本數據文件](sample_data/sample.xlsx)的格式或問卷的編碼方式。

## 項目結構

```
├── data/                 # 數據文件
├── img/                  # 簡報使用的圖片
├── output/               # 生成的簡報
├── src/
│   ├── processors/       # 數據處理模組
│   ├── data_reader.py    # 數據讀取與清理
│   ├── data_validator.py # 數據驗證規則
│   ├── ppt_generator.py  # PowerPoint 生成
│   ├── presentation_generator.py # 主簡報生成器
│   └── streamlit.py      # Streamlit 網頁界面
├── requirements.txt      # Python 依賴項
└── README.md
```
