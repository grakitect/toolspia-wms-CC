/**
 * 정보 가리기(데모) 모드 공용 마스킹 유틸.
 * 화면 표시용 가짜 값만 만들어내며 실제 데이터(서버/DB)는 절대 건드리지 않는다.
 * - 같은 원본 문자열은 항상 같은 가짜 값으로 치환된다(해시 기반, 랜덤 아님).
 * - 서로 다른 원본은 서로 다른 가짜 값으로 치환된다(같은 세션 내 중복 없음, 해시 충돌 시 다음 후보로 이동).
 * - 상품명은 같은 행의 실제 "카테고리"와 어울리는 단어 풀에서 생성해 카테고리와 이름이 상식적으로 맞게 한다.
 * index.html(app.js)과 barcode-print.html이 동일한 로직을 쓰도록 이 파일을 공유한다.
 */
(function (global) {
  function hashStr(s) {
    let h = 0;
    for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0;
    return h;
  }

  // 카테고리 키워드 → 어울리는 형용사/명사/단위 풀. 조합(형용사×명사×단위)으로 충분히 큰 풀을 만들어
  // 상품 수가 많아도 서로 다른 이름을 배정할 수 있게 한다. 식품 위주가 아니라 생활용품 전반으로 구성.
  const CATEGORY_BANKS = [
    { keywords: ["공구", "생품"], adjectives: ["정밀", "다용도", "휴대용", "경량", "고강도", "접이식", "자동", "수동", "미니", "대형"], nouns: ["드라이버세트", "니퍼", "펜치", "줄자", "글루건", "육각렌치세트", "커터", "해머", "스패너", "전동드릴비트", "가스토치"], units: ["", "1세트", "2p", "3p", "5p", "10p"] },
    { keywords: ["원예"], adjectives: ["원예용", "정원용", "다육", "텃밭", "실내용", "야외용", "다용도"], nouns: ["모종삽", "전지가위", "화분", "분갈이흙", "식물영양제", "원예장갑", "물뿌리개", "지지대", "화분받침"], units: ["", "1개", "2p", "1L", "500g"] },
    { keywords: ["수전", "샤워"], adjectives: ["절수형", "회전형", "분리형", "고정형", "스텐", "다단계"], nouns: ["샤워기헤드", "수도꼭지", "연장호스", "연결어댑터", "샤워기거치대", "수전연결관"], units: ["", "1개", "2m", "4m"] },
    { keywords: ["필터", "필조"], adjectives: ["정수", "교체용", "다단계", "활성탄", "고성능"], nouns: ["샤워필터", "정수필터", "리필필터", "필터카트리지", "필터세트", "필터조립품"], units: ["", "1개", "3p", "4p", "6p"] },
    { keywords: ["접착", "실란트", "실리콘"], adjectives: ["강력", "속건", "다목적", "방수", "고정력"], nouns: ["순간접착제", "실란트", "실리콘마감재", "접착테이프", "본드", "누수방지테이프"], units: ["", "1개", "300ml", "2p"] },
    { keywords: ["철물"], adjectives: ["고정형", "다용도", "스텐", "경량", "매립형"], nouns: ["선반브라켓", "경첩", "타격앙카", "액자걸이", "볼트너트세트", "몰딩고정클립"], units: ["", "1개", "2p", "12p"] },
    { keywords: ["케미"], adjectives: ["다목적", "강력", "친환경", "고농축"], nouns: ["녹제거제", "다목적클리너", "접착제거제", "세정액", "방청윤활제"], units: ["", "1개", "20g", "150ml"] },
    { keywords: ["완구", "장난감"], adjectives: ["교육용", "유아용", "조립식", "미니", "대형"], nouns: ["블록세트", "퍼즐", "인형", "자동차완구", "역할놀이세트"], units: ["", "1개", "1세트"] },
    { keywords: ["문구", "학용품"], adjectives: ["기본형", "휴대용", "다용도", "심플"], nouns: ["볼펜세트", "노트", "클립보드", "가위", "파일홀더", "연필세트"], units: ["", "1개", "2p", "10p"] },
    { keywords: ["주방"], adjectives: ["다용도", "실리콘", "스테인리스", "휴대용", "대용량"], nouns: ["밀폐용기", "주방장갑", "조리도구세트", "도마", "계량컵", "주방집게"], units: ["", "1개", "1세트", "2p"] },
    { keywords: ["욕실"], adjectives: ["미끄럼방지", "방수", "흡착형", "다용도"], nouns: ["욕실매트", "샤워커튼", "비누받침", "욕실선반", "칫솔꽂이"], units: ["", "1개", "1세트"] },
    { keywords: ["청소"], adjectives: ["다용도", "회전형", "물걸레", "정전기방지"], nouns: ["청소솔", "밀대", "먼지제거포", "청소용스퀴지", "걸레세트"], units: ["", "1개", "1세트", "10매"] },
    { keywords: ["수납"], adjectives: ["다용도", "접이식", "투명", "적층형"], nouns: ["수납정리함", "옷걸이", "서랍정리대", "행거", "바구니"], units: ["", "1개", "1세트"] },
    { keywords: ["조명"], adjectives: ["LED", "무드", "센서형", "충전식"], nouns: ["조명등", "무드등", "센서등", "손전등", "스탠드조명"], units: ["", "1개"] },
    { keywords: ["전자", "전기", "멀티"], adjectives: ["휴대용", "충전식", "무선", "미니", "고용량"], nouns: ["보조배터리", "usb허브", "충전케이블", "미니선풍기", "전자저울", "멀티탭"], units: ["", "1개", "2구", "3구"] },
    { keywords: ["반려동물", "펫"], adjectives: ["반려동물용", "휴대용", "세척가능"], nouns: ["펫방석", "급수기", "장난감", "이동가방", "브러시"], units: ["", "1개"] },
    { keywords: ["뷰티", "미용"], adjectives: ["휴대용", "다용도", "저자극"], nouns: ["헤어브러시", "메이크업퍼프", "손톱깎이세트", "미니거울"], units: ["", "1개", "1세트"] },
    { keywords: ["자동차", "차량"], adjectives: ["차량용", "휴대용", "다용도"], nouns: ["방향제", "핸들커버", "차량용거치대", "세차타월"], units: ["", "1개", "1세트"] },
    { keywords: ["스포츠", "레저", "캠핑"], adjectives: ["휴대용", "경량", "야외용", "접이식"], nouns: ["캠핑의자", "돗자리", "보온병", "손선풍기", "우산"], units: ["", "1개"] },
    { keywords: ["의류", "패션", "잡화"], adjectives: ["기본형", "사계절용", "심플"], nouns: ["양말", "장갑", "손수건", "모자", "머플러"], units: ["", "1개", "1세트"] },
  ];
  const GENERIC_BANK = { adjectives: ["기본형", "다용도", "휴대용", "실용적인", "심플한", "고급형", "가정용", "인기", "베스트", "신형"], nouns: ["생활용품", "정리용품", "수납케이스", "생활잡화", "다용도홀더", "일상소품", "생활도구", "편의용품"], units: ["", "1개", "1세트", "2p"] };
  const GENERIC_BANK_EN = { adjectives: ["Basic", "Compact", "Premium", "Everyday", "Multi-use", "Portable"], nouns: ["Storage Box", "Household Item", "Organizer", "Utility Tool", "Home Goods"], units: ["", "Set", "Pack"] };

  const COMPANY_ADJ_KO = ["가나", "다라", "마바", "사아", "자차", "카타", "파하", "온빛", "미르", "누리", "별빛", "한들", "새벽", "은하", "solar"];
  const COMPANY_NOUN_KO = ["상사", "유통", "물류", "상회", "무역", "산업", "통상", "인터내셔널"];
  const COMPANY_ADJ_EN = ["Acme", "Northwind", "Globex", "Umbrella", "Initech", "Hooli", "Wayne", "Stark", "Sample", "Demo"];
  const COMPANY_NOUN_EN = ["Corp", "Trading", "Logistics", "Co", "Ltd", "Inc", "LLC", "Imports"];

  const KEEP_AS_IS = new Set(["판매중", "단종", "판매중단", "단일", "다중", "-"]);

  function buildCombos(bank) {
    const out = [];
    for (const adj of bank.adjectives) {
      for (const noun of bank.nouns) {
        for (const unit of bank.units) {
          out.push(unit ? `${adj} ${noun} ${unit}` : `${adj} ${noun}`);
        }
      }
    }
    return out;
  }

  function pickProductBank(categoryHint) {
    const hint = String(categoryHint || "").trim();
    if (hint) {
      for (const bank of CATEGORY_BANKS) {
        if (bank.keywords.some((kw) => hint.includes(kw))) return bank;
      }
    }
    return null;
  }

  // 세션 동안 유지되는 배정 캐시(원본→가짜) 및 사용된 가짜 값 집합(카테고리 풀별로 분리해
  // "충전식 usb허브"가 전자 카테고리에서만 쓰이는 식으로 자연스럽게 겹치지 않게 한다).
  const productCache = new Map(); // key: categoryHint + "|" + text -> fake
  const usedByPool = new Map(); // key: pool identity(문자열) -> Set(used fake)
  const companyCache = new Map();
  const usedCompanyNames = new Set();

  function poolKeyFor(bank, hasKorean) {
    if (!bank) return hasKorean ? "generic-ko" : "generic-en";
    return bank.keywords[0] + (hasKorean ? "-ko" : "-en");
  }

  function assignUnique(pool, cacheKey, cache, usedSet) {
    if (cache.has(cacheKey)) return cache.get(cacheKey);
    const n = pool.length;
    const start = hashStr(cacheKey) % n;
    for (let i = 0; i < n; i++) {
      const candidate = pool[(start + i) % n];
      if (!usedSet.has(candidate)) {
        usedSet.add(candidate);
        cache.set(cacheKey, candidate);
        return candidate;
      }
    }
    // 풀이 다 소진된 경우(상품 수가 매우 많을 때)에도 유일성을 보장하기 위해 번호를 붙인다.
    const fallback = `${pool[start]} ${(hashStr(cacheKey + "#") % 900) + 100}`;
    usedSet.add(fallback);
    cache.set(cacheKey, fallback);
    return fallback;
  }

  function randomizeProductName(text, categoryHint) {
    const s = String(text ?? "");
    const hasKorean = /[가-힣]/.test(s);
    const bank = pickProductBank(categoryHint);
    const effectiveBank = bank || (hasKorean ? GENERIC_BANK : GENERIC_BANK_EN);
    const poolKey = poolKeyFor(bank, hasKorean);
    if (!usedByPool.has(poolKey)) usedByPool.set(poolKey, new Set());
    const pool = buildCombos(effectiveBank);
    const cacheKey = poolKey + "|" + s;
    return assignUnique(pool, cacheKey, productCache, usedByPool.get(poolKey));
  }

  function randomizeCompanyName(text) {
    const s = String(text ?? "");
    const hasKorean = /[가-힣]/.test(s);
    const pool = [];
    const adjs = hasKorean ? COMPANY_ADJ_KO : COMPANY_ADJ_EN;
    const nouns = hasKorean ? COMPANY_NOUN_KO : COMPANY_NOUN_EN;
    for (const a of adjs) for (const n of nouns) pool.push(hasKorean ? `${a}${n}` : `${a} ${n}`);
    return assignUnique(pool, s, companyCache, usedCompanyNames);
  }

  function columnCategory(label) {
    const l = String(label || "").trim();
    if (l.includes("품목명") || l.includes("상품명")) return "product";
    if (l === "판매처" || l === "구매처") return "company";
    return null;
  }

  function randomizeText(text, category, categoryHint) {
    const s = String(text ?? "");
    if (!s.trim()) return s;
    if (KEEP_AS_IS.has(s.trim())) return s;
    if (category === "product") return randomizeProductName(s, categoryHint);
    if (category === "company") return randomizeCompanyName(s);
    if (/^-?\d[\d,.]*%?$/.test(s.trim())) {
      return s.replace(/\d/g, (m, idx) => String(hashStr(s + ":" + idx) % 10));
    }
    const hasLatinDigitMix = /[A-Za-z]/.test(s) && /\d/.test(s);
    if (hasLatinDigitMix) {
      return s
        .replace(/[A-Za-z]/g, (m, idx) => String.fromCharCode(97 + (hashStr(s + ":a:" + idx) % 26)))
        .replace(/\d/g, (m, idx) => String(hashStr(s + ":d:" + idx) % 10));
    }
    return randomizeCompanyName(s);
  }
  function randomizeList(list, category, categoryHint) {
    if (!Array.isArray(list)) return list;
    return list.map((v) => randomizeText(v, category, categoryHint));
  }

  global.WmsDemoMask = { randomizeText, randomizeList, columnCategory, hashStr, STORAGE_KEY: "wms:demoMode" };
})(window);
