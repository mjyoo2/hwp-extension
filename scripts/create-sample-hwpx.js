const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function createSampleHwpx() {
  const zip = new JSZip();

  const versionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:HWPApplicationSetting xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:ApplicationVersion>11.0.0.0</hh:ApplicationVersion>
</hh:HWPApplicationSetting>`;

  const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"
         xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">
  <hh:docInfo>
    <hh:title>HWPX Editor 테스트 문서</hh:title>
    <hh:creator>HWPX Editor</hh:creator>
    <hh:createdDate>2024-01-01T00:00:00</hh:createdDate>
    <hh:modifiedDate>2024-01-01T00:00:00</hh:modifiedDate>
    <hh:description>VSCode HWPX Editor 테스트용 샘플 문서</hh:description>
  </hh:docInfo>
  <hh:refList>
    <hh:fontfaces itemCnt="2">
      <hh:fontface lang="HANGUL" fontCnt="1">
        <hh:font id="0" face="맑은 고딕" type="TTF"/>
      </hh:fontface>
      <hh:fontface lang="LATIN" fontCnt="1">
        <hh:font id="0" face="맑은 고딕" type="TTF"/>
      </hh:fontface>
    </hh:fontfaces>
    <hh:charProperties itemCnt="3">
      <hh:charPr id="0" height="1000" textColor="#000000">
        <hh:fontRef hangul="0" latin="0"/>
      </hh:charPr>
      <hh:charPr id="1" height="1400" textColor="#000000" bold="1">
        <hh:fontRef hangul="0" latin="0"/>
      </hh:charPr>
      <hh:charPr id="2" height="1000" textColor="#0000FF">
        <hh:fontRef hangul="0" latin="0"/>
        <hh:underline type="BOTTOM" shape="SOLID" color="#0000FF"/>
      </hh:charPr>
    </hh:charProperties>
    <hh:paraProperties itemCnt="3">
      <hh:paraPr id="0">
        <hh:align horizontal="LEFT"/>
      </hh:paraPr>
      <hh:paraPr id="1">
        <hh:align horizontal="CENTER"/>
      </hh:paraPr>
      <hh:paraPr id="2">
        <hh:align horizontal="JUSTIFY"/>
        <hh:lineSpacing type="PERCENT" value="160"/>
      </hh:paraPr>
    </hh:paraProperties>
  </hh:refList>
</hh:head>`;

  const section0Xml = `<?xml version="1.0" encoding="UTF-8"?>
<hp:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
        xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section">
  <hp:p paraPrIDRef="1">
    <hp:run charPrIDRef="1">
      <hp:t>HWPX Editor 테스트 문서</hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t></hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="2">
    <hp:run charPrIDRef="0">
      <hp:t>이 문서는 VSCode에서 HWPX 파일을 편집하기 위한 테스트용 샘플입니다. 리브레오피스 수준의 편집 기능을 지원합니다. 텍스트 서식, 단락 정렬, 표, 이미지 등 다양한 기능을 테스트할 수 있습니다.</hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t></hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="1">
      <hp:t>지원 기능:</hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t>• 텍스트 서식: 굵게, 기울임, 밑줄, 취소선</hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t>• 글꼴 설정: 글꼴, 크기, 색상</hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t>• 단락 서식: 정렬, 줄간격, 들여쓰기</hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t>• 표 편집: 행 추가/삭제, 셀 편집</hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t>• 이미지 표시</hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t></hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="1">
      <hp:t>샘플 표:</hp:t>
    </hp:run>
  </hp:p>
  <hp:tbl>
    <hp:tr>
      <hp:tc>
        <hp:p><hp:run><hp:t>항목</hp:t></hp:run></hp:p>
      </hp:tc>
      <hp:tc>
        <hp:p><hp:run><hp:t>설명</hp:t></hp:run></hp:p>
      </hp:tc>
      <hp:tc>
        <hp:p><hp:run><hp:t>비고</hp:t></hp:run></hp:p>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc>
        <hp:p><hp:run><hp:t>HWPX</hp:t></hp:run></hp:p>
      </hp:tc>
      <hp:tc>
        <hp:p><hp:run><hp:t>한글 문서 XML 형식</hp:t></hp:run></hp:p>
      </hp:tc>
      <hp:tc>
        <hp:p><hp:run><hp:t>KS X 6101 표준</hp:t></hp:run></hp:p>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc>
        <hp:p><hp:run><hp:t>VSCode</hp:t></hp:run></hp:p>
      </hp:tc>
      <hp:tc>
        <hp:p><hp:run><hp:t>코드 에디터</hp:t></hp:run></hp:p>
      </hp:tc>
      <hp:tc>
        <hp:p><hp:run><hp:t>Microsoft</hp:t></hp:run></hp:p>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t></hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t>────────────────────────────────────────</hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="1">
      <hp:t>그래픽 요소 테스트:</hp:t>
    </hp:run>
  </hp:p>
  <hp:line startX="0" startY="0" endX="50000" endY="0" lineColor="#3366CC" lineWidth="200"/>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t>위는 파란색 구분선입니다.</hp:t>
    </hp:run>
  </hp:p>
  <hp:rect x="0" y="0" width="20000" height="10000" fillColor="#FFFFCC" lineColor="#CC9900" lineWidth="100"/>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t>위는 노란색 사각형입니다.</hp:t>
    </hp:run>
  </hp:p>
  <hp:ellipse centerX="5000" centerY="5000" radiusX="5000" radiusY="3000" fillColor="#CCFFCC" lineColor="#339933" lineWidth="100"/>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t>위는 초록색 타원입니다.</hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t>━━━━━━━━━━━━━━━━━━━━</hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t></hp:t>
    </hp:run>
  </hp:p>
  <hp:p paraPrIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:t>각 단락을 클릭하여 편집할 수 있습니다. Ctrl+B(굵게), Ctrl+I(기울임), Ctrl+U(밑줄) 단축키를 사용할 수 있습니다.</hp:t>
    </hp:run>
  </hp:p>
</hp:sec>`;

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`;

  zip.file('version.xml', versionXml);
  zip.file('Contents/header.xml', headerXml);
  zip.file('Contents/section0.xml', section0Xml);
  zip.file('[Content_Types].xml', contentTypesXml);

  const outputPath = path.join(__dirname, '..', 'test-files', 'sample.hwpx');
  const content = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(outputPath, content);

  console.log('Sample HWPX file created:', outputPath);
}

createSampleHwpx().catch(console.error);
