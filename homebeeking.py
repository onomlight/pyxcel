import pandas as pd
from tkinter import messagebox
# 엑셀 파일 경로
file_path = 'bigtest.xlsx'
output_file_path = 'new__새로운_파일.xlsx'

# 엑셀 파일 불러오기
df = pd.read_excel(file_path)

# 'O'열 초기화
df['O'] = ''

# 조건과 매칭될 값을 정의한 딕셔너리
mapping_dict = {
('1+1 살림백서 천연 지수81% 샴푸 500ml 약산성 퍼퓸 비듬 향기좋은 사춘기 청소년 초등학생 지성 두피 정수리냄새 대용량 오푼티아 밤부', '향:유칼립투스향 샴푸 500ml 1+1-1개'): '♡4W_OPS0006_샴푸 (유칼립투스) 500ml 2개입',
('살림백서 주방세제 리필형 1L 3개입 대용량 천연유래98% 리필 과일세척', '향:살림백서 주방세제 리필 1Lx3개(청귤)-2개'): '■4WC_SB0010_주방세제 리필형 1L (청귤향) 3개입',
('살림백서 주방세제 리필형 1L 3개입 대용량 천연유래98% 리필 과일세척', '향:살림백서 주방세제 리필 1Lx3개(청귤)-1개'): '■4WC_SB0010_주방세제 리필형 1L (청귤향) 3개입',
('살림백서 주방세제 리필형 1L 3개입 대용량 천연유래98% 리필 과일세척', '향:살림백서 주방세제 리필 1Lx3개(자몽)-1개'): '■4WC_SB0007_주방세제 리필형 1L (자몽향) 3개입',
('1+1 살림백서 천연 지수81% 샴푸 500ml 약산성 퍼퓸 비듬 향기좋은 사춘기 청소년 초등학생 지성 두피 정수리냄새 대용량 오푼티아 밤부', '향:베이비파우더향 샴푸 500ml 1+1-1개'): '♡4W_OPS0001_샴푸 (베이비파우더) 500ml 2개입',
('1+1 살림백서 대용량 바디로션 라이스&허브 1000ml 향좋은 고보습 촉촉한 끈적임없는', '향:RH2.그린포레스트 바디로션 2개-1개'): '향:RH2.그린포레스트 바디로션 2개-1개',
('1+1 살림백서 라이스 앤 허브 샴푸 1L 약산성 퍼퓸 향기좋은 정수리 냄새 청소년 사춘기', '향:RH1.딥그린시더우드 샴푸 (1+1)-1개'): '♡4W_RH_OPR002_R허브 샴푸 딥그린시더우드 1L 2개입',
('1+1 살림백서 라이스 앤 허브 바디워시 1L 대용량 퍼퓸 향좋은 샤워젤 바디클렌져', '향:RH1.딥그린시더우드 바디워시 (1+1)-1개'): '♡4W_RH_OPR010_R허브 바디워시 딥그린시더우드 1L 2개입',
('1+1 살림백서 디퓨저 라이스&허브 200ml 아로마 실내 방향제 화장실 인테리어 대용량', '향:4)1+1 살림백서 디퓨저 포레스트가든-1개'): '♡4W_RH_SB0115_R허브 디퓨저 (포레스트가든) 200ml 2개입',
('살림백서 소프트 퍼퓸 건조기시트 80매 건조기섬유유연제 드라이시트 3종향', '향:3)퍼퓸 건조기 섬유유연제 80매(코트니향)-1개'): '♥4MC_SB0076_건조기 섬유유연제 (코트니향) 2박스',
('살림백서 딥퍼퓸 섬유유연제 1L X 3개입 5종향 실내건조 향기좋은 대용량', '향:2.딥퍼퓸 섬유유연제 베이비파우더 1L 3개-1개'): '♥3MC_SB0129_딥퍼퓸 섬유유연제 (베이비파우더) 1L 3개입',
('살림백서 소프트 퍼퓸 건조기시트 80매 건조기섬유유연제 드라이시트 3종향', '향:2)퍼퓸 건조기 섬유유연제80매(러브미향)-1개'): '♥3M_SB0074_건조기 섬유유연제 (러브미향) 2박스',
('1+1 살림백서 디퓨저 라이스&허브 200ml 아로마 실내 방향제 화장실 인테리어 대용량', '향:2)1+1 살림백서 디퓨저 클린솝-1개'): '♡4W_RH_SB0116_R허브 디퓨저 (클린솝) 200ml 2개입',
('살림백서 딥퍼퓸 섬유유연제 1L X 3개입 5종향 실내건조 향기좋은 대용량', '향:1.딥퍼퓸 섬유유연제 화이트머스크 1L 3개-1개'): '■3MC_SB0128_딥퍼퓸 섬유유연제 (화이트머스크) 1L 3개입',
('1+1 살림백서 디퓨저 라이스&허브 200ml 아로마 실내 방향제 화장실 인테리어 대용량', '향:1)1+1 살림백서 디퓨저 일랑일랑-1개'): '♡4W_RH_SB0118_R허브 디퓨저 (일랑일랑) 200ml 2개입',
('살림백서 소프트 퍼퓸 건조기 섬유유연제 40매 3종향', '향 선택)=2)퍼퓸 건조기 섬유유연제 40매(러브미향)'): '♥4MC_SB0073_건조기 섬유유연제 (러브미향) 1박스',
('살림백서 소프트 퍼퓸 건조기 섬유유연제 40매 3종향', '향 선택)=1)퍼퓸 건조기 섬유유연제 40매(블랑향)'): '♥4MC_SB0071_건조기 섬유유연제 (블랑향) 1박스',
('살림백서 딥퍼퓸 섬유유연제 1L x 3개입 5종향', '제품 선택=딥퍼퓸 섬유유연제 1L 3개, 향기 선택=화이트머스크'): '♥3MC_SB0128_딥퍼퓸 섬유유연제 (화이트머스크) 1L 3개입',
('1+1 살림백서 곰팡이제거제 젤 벽 벽지 화장실 베란다 실리콘 창틀 욕실 결로 곰팡이', '옵션선택):1+1+1 살림백서 바르는 곰팡이제거제-1개'): '♥3MC_SB0016_바르는 곰팡이제거제 150g 3개입',
('1+1 마미바티 아기세제 아기섬유유연제', '선택1: 01_1+1 마미바티 아기세제1800ml'): '♥3M_OMVA1002_마미바티 아기세제 2개입',
('1+1 마미바티 아기세제 1.8L 유아 세탁세제 신생아', '선택1: 01_1+1 마미바티 아기세제1800ml'): '♥3M_OMVA1002_마미바티 아기세제 2개입',
('살림백서 실내건조 세탁세제 2.7L x 4개입 대용량 액체세제 드럼용 일반용', '상품선택=2)실내건조 세탁세제 2.7L 4개입(드럼용)'): '■4WC_SB0101_실내 건조용 세탁세제 (드럼) 2.7L 2개입',
('살림백서 실내건조 세탁세제 2.7L x 4개입 대용량 액체세제 드럼용 일반용', '상품선택=1)실내건조 세탁세제 2.7L 4개입(일반용)'): '■4WC_SB0100_실내 건조용 세탁세제 (일반) 2.7L 4개입',
('살림백서 대용량 1종 주방세제 라이스&허브 4L+750ml', '살림백서 라이스&허브 주방세제 4L+750ml=1)주방세제 석류향 4L+750ml'): '■4W_SB0106_R허브 주방세제 (석류) 4L 1개입+750ml 1개입',
('살림백서 에어프라이어 종이호일 90매 (중형/대형) 기름종이 에어프라이 오븐 원형', '사이즈:대형 90매-2개 (+2000원)'): '■1M_SBI001_종이호일 대형 90매',
('살림백서 에어프라이어 종이호일 90매 (중형/대형) 기름종이 에어프라이 오븐 원형', '사이즈:대형 90매-1개 (+1000원)'): '■1M_SBI001_종이호일 대형 90매',
('살림백서 하루세번 구취케어 치약 100g 3개입 입냄새 구취제거 잇몸', '사이즈:01.하루세번 구취케어 치약 100g 3개입-1개'): '♥3M_SB0154_구취케어 치약 3개입',
('베이킹소다 대용량 세탁세제 2.5L 4개입 중성세제 얼룩제거 냄새제거 일반 드럼 세탁기겸용', '베이킹소다 세탁세제 2.5L 4개'): '■4WC_OCOA1003_베이킹소다 세탁세제 2.5L 4개입 ♣',
('1+1 살림백서 오푼티아&밤부 바디로션 500ml', '바디로션 향 선택1=03.바디로션 베베머스크, 바디로션 향 선택2=03.바디로션 베베머스크'): '♡4W_OPBL004_바디로션 베베머스크향 500ml 2개입',
('1+1 살림백서 오푼티아&밤부 바디로션 500ml', '바디로션 향 선택1=02.바디로션 화이트머스크, 바디로션 향 선택2=04.바디로션 체리블라썸'): '♡4W_OPBL106_바디로션 화이트머스크+체리블라썸 500ml',
('천연유래 아기세제 1.8L 2개', '마미바티 아기세제 2개입'): '♥3M_OMVA1002_마미바티 아기세제 2개입',
('아기 섬유유연제 1.8L 2개', '마미바티 아기섬유유연제 1.8L 2개'): '♥3M_OMVA1005_마미바티 섬유유연제 2개입',
('살림백서 베이킹소다 3kg 리필형1+1, 3000g, 1개', '1개 3000g'): '■4W_SBF002_베이킹소다 3kg 2개입',
('1+1 살림백서 천연유래98% 세탁세제', '1+1 살림백서 천연유래98% 세탁세제'): '♥4MC_SBA001_세탁세제 2L (본품) 2개입',
('1+1 살림백서 섬유유연제 2L 2개', '1+1 살림백서 섬유유연제 2L 2개'): '♥4MC_SBB001_섬유유연제 2L (본품) 2개입',
('1+1 살림백서 엑티브B7 맥주효모 앤 비오틴 탈모완화 샴푸 1000ml 대용량 지성 두피 남자 여자 약산성', '1+1 살림백서 엑티브B7 맥주효모 앤 비오틴 탈모완화 샴푸 1000ml 대용량 지성 두피 남자 여자 약산성'): '■4WC_OPP004_탈모샴푸 1L 2개입',
('살림백서 2단계 초 미세모 칫솔 딥그래이쉬 20개입 초극세모 임산부 대용량', '살림백서 2단계 초 미세모 칫솔 딥그래이쉬 20개입 초극세모 임산부 대용량'): '♥4MC_SB0126_딥그레이쉬 칫솔 20개입 1박스',
('살림백서 먼지털이 스틱+리필10개입 먼지떨이 세트 떨이개 제거기 털이개', '살림백서 먼지털이 스틱+리필10개입 먼지떨이 세트 떨이개 제거기 털이개'): '♥4MC_SB0121_먼지떨이 스틱+리필10개입 SET',
('마미바티 아기 섬유유연제 1+1 (단7일 행사가)', '마미바티 아기 섬유유연제 1+1 (단7일 행사가)'): '♥3M_OMVA1005_마미바티 섬유유연제 2개입',
('살림백서 울세제 중성세제 리필형 1L 3개입', '살림백서 울세제 중성세제 리필형 1L 3개입'): '■4WC_SB0013_울세제 리필형 1L 3개입',
('살림백서 실내건조 세탁세제 2.7L 4개입(드럼용) 대용량 액체세제', '살림백서 실내건조 세탁세제 2.7L 4개입(드럼용) 대용량 액체세제'): '■4WC_SB0102_실내 건조용 세탁세제 (드럼) 2.7L 4개입',
('살림백서 변기세정제 (40g X 10개)', '살림백서 변기세정제 (40g X 10개)'): '♥4MC_SB0050_변기세정제 1박스',
('(단하루!) 살림백서 베이킹소다 3kg 대용량 리필형 1+1', '(단하루!) 살림백서 베이킹소다 3kg 대용량 리필형 1+1'): '■4W_SBF002_베이킹소다 3kg 2개입',
('1+1 살림백서 식기세척기 세제 분말형 1kg 식세기 sk lg 밀레 삼성 호환', '1+1 살림백서 식기세척기 세제 분말형 1kg 식세기 sk lg 밀레 삼성 호환'): '♥3MC_SB0039_식세기 세제 1kg (분말) 2개입',
('살림백서 실내건조 세탁세제 2.7L 2개입(일반용) 대용량 액체세제', '살림백서 실내건조 세탁세제 2.7L 2개입(일반용) 대용량 액체세제'): '■4WC_SB0099_실내 건조용 세탁세제 (일반) 2.7L 2개입',
('1+1 살림백서 세탁조 클리너 (총 8회분)', '1+1 살림백서 세탁조 클리너 (총 8회분)'): '♥5M_SB0003_세탁조 클리너 2박스',
('1+1 살림백서 울세제 1L', '1+1 살림백서 울세제 1L'): '♥3MC_SB0012_울세제 1L 2개입',
('1+1 살림백서 액체세탁세제 2L', '1+1 살림백서 액체세탁세제 2L'): '♥4MC_SBA001_세탁세제 2L (본품) 2개입',
('1+1 살림백서 울세제 중성세제 1L 홈드라이 세탁', '1+1 살림백서 울세제 중성세제 1L 홈드라이 세탁'): '♥3MC_SB0012_울세제 1L 2개입',
('살림백서 천연유래98% 액체 세탁세제 2L 6개입 중성세제 드럼용 일반용 겸용', '살림백서 천연유래98% 액체 세탁세제 2L 6개입 중성세제 드럼용 일반용 겸용'): 'SB_A_3/ 세탁세제 2L (리필) 6개입',
('살림백서 뽑아쓰는 키친타올 100매 9팩', '살림백서 뽑아쓰는 키친타올 100매 9팩'): 'SB_3PL_2/ 살림백서 (뽑)키친타월 1BOX',
('살림백서 보들보들 각티슈 미용티슈 250매 6개입', '살림백서 보들보들 각티슈 미용티슈 250매 6개입'): '3PL_3/ 살림백서 미용티슈 1BOX',
('살림백서 베이비&키즈 아기세제 아기세탁세제 클린솝 1950ml 2개입', '살림백서 베이비&키즈 아기세제 아기세탁세제 클린솝 1950ml 2개입'): '♥3M_BA0007_아기유연제 1950ml 2개입',
('살림백서 찌든때 클리너 240g 2개입 스테인레스 연마제거제 탄냄비닦는법 텀블러 세척', '살림백서 찌든때 클리너 240g 2개입 스테인레스 연마제거제 탄냄비닦는법 텀블러 세척'): '♥4MC_SB0019_찌든때 클리너 1박스',
('살림백서 변기세정제 40g 10개입', '살림백서 변기세정제 40g 10개입'): '♥4MC_SB0050_변기세정제 1박스',
('마미바티 과탄산소다 2kg 산소계표백제', '마미바티 과탄산소다 2kg 산소계표백제'): '♥3M_OMVA1008_마미바티 과탄산 용기 1개입',
('1+1 살림백서 배수구 클리너 (총 8회분)', '1+1 살림백서 배수구 클리너 (총 8회분)'): '♥5M_SB0001_배수구 클리너 2박스',
('1+1 살림백서 배수구클리너 150g 4개 (총 8회분)', '1+1 살림백서 배수구클리너 150g 4개 (총 8회분)'): '♥5M_SB0001_배수구 클리너 2박스',
('(단7일! 행사가) 마미바티 아기 젖병세정제 1 + 1  500ml', '(단7일! 행사가) 마미바티 아기 젖병세정제 1 + 1  500ml'): '♥4MC_OMVB1002_마미바티 주방세제 2개입',
('1+1 살림백서 섬유유연제 2L', '1+1 살림백서 섬유유연제 2L'): '♥4MC_SBB001_섬유유연제 2L (본품) 2개입',
('살림백서 구연산 1kg 리필형 1+1', '살림백서 구연산 1kg 리필형 1+1'): '■4WC_SBG002_구연산 1kg 2개입',
('1+1 살림백서 뿌리는 신발탈취제 200ml 신발장 신발냄새제거제', '1+1 살림백서 뿌리는 신발탈취제 200ml 신발장 신발냄새제거제'): '■4WC_SB0052_슈즈케어 2박스',
('살림백서 제습제 습기제거제 520mlx24개', '살림백서 제습제 습기제거제 520mlx24개'): 'SB_3PL_1/ 살림백서 제습제 1BOX',
('살림백서 발각질 제거제 250ml 풋 필링 스프레이 타입', '살림백서 발각질 제거제 250ml 풋 필링 스프레이 타입'): '♡4W_OPP008_발각질 제거제 250ml 1개입',
('1+1 살림백서 세탁조 클리너 드럼세탁기 청소 통돌이 통세척 (총 8회분)', '1+1 살림백서 세탁조 클리너 드럼세탁기 청소 통돌이 통세척 (총 8회분)'): '♥5M_SB0003_세탁조 클리너 2박스',
('1+1 살림백서 과탄산소다 3kg 대용량 리필형', '1+1 살림백서 과탄산소다 3kg 대용량 리필형'): '■4WC_SBE002_과탄산소다 3kg 2개입',
('살림백서 흰옷 표백제 1kg X 2 EA', '살림백서 흰옷 표백제 1kg X 2 EA'): '살림백서 흰옷 표백제 1kg X 2 EA',
('살림백서 옷걸이형 제습제 옷장형 습기제거제 150g 10개입', '살림백서 옷걸이형 제습제 옷장형 습기제거제 150g 10개입'): '■4MC_SB0084_옷걸이형 제습제 (비드형) 10개입',
('1+1 살림백서 주방세제 청귤향 1L(펌프포함)', '1+1 살림백서 주방세제 청귤향 1L(펌프포함)'): '♥3MC_SB0009_주방세제 1L (청귤향) 2개입',
('메르헨트 퍼퓸 드 디퓨저 200ml 2개입 일랑일랑', '메르헨트 퍼퓸 드 디퓨저 200ml 2개입 일랑일랑'): '메르헨트 퍼퓸 드 디퓨저 200ml 2개입 일랑일랑',
('1+1 살림백서 산소계 표백제 1kg 흰옷 얼룩제거', '1+1 살림백서 산소계 표백제 1kg 흰옷 얼룩제거'): '■4WC_SBH001_흰옷표백제 1kg 2개입',
('1+1 살림백서 세탁조 클리너 드럼세탁기 청소 통돌이 통세척 8회분 (150g x 8개)', '1+1 살림백서 세탁조 클리너 드럼세탁기 청소 통돌이 통세척 8회분 (150g x 8개)'): '♥5M_SB0003_세탁조 클리너 2박스',
('메르헨트 퍼퓸 드 섬유향수 룸스프레이, 소프트코튼, 250ml, 2개', '메르헨트 퍼퓸 드 섬유향수 룸스프레이, 소프트코튼, 250ml, 2개'): '메르헨트 퍼퓸 드 섬유향수 룸스프레이, 소프트코튼, 250ml, 2개',
('살림백서 식기세척기 린스 헹굼보조제 가정용 액상형 1L sk lg 밀레 삼성 호환', '살림백서 식기세척기 린스 헹굼보조제 가정용 액상형 1L sk lg 밀레 삼성 호환'): '♥3M_SB0040_식세기 헹굼보조제 1L 1개입',
('1+1 살림백서 과탄산소다 3kg (총 6kg 대용량)', '1+1 살림백서 과탄산소다 3kg (총 6kg 대용량)'): '■4WC_SBE002_과탄산소다 3kg 2개입',
('살림백서 만능 얼룩제거제 리필형 1L 2개입', '살림백서 만능 얼룩제거제 리필형 1L 2개입'): '♥3MC_SB0020_얼룩제거제 500ml 2개입',
('살림백서 제습제 520ml 24개입 1box 옷장 화장실 신발장', '살림백서 제습제 520ml 24개입 1box 옷장 화장실 신발장'): 'SB_3PL_1/ 살림백서 제습제 1BOX',
('1+1 살림백서 탈모 샴푸 엑티브B7 맥주효모 앤 비오틴 1000ml', '1+1 살림백서 탈모 샴푸 엑티브B7 맥주효모 앤 비오틴 1000ml'): '■4WC_OPP004_탈모샴푸 1L 2개입',
('1+1 살림백서 화장실 욕실세정제 800ml 청소세제 곰팡이제거 소독제', '1+1 살림백서 화장실 욕실세정제 800ml 청소세제 곰팡이제거 소독제'): '♥4W_SB0082_욕실세정제 800ml 2개입',
('피부백서 순면 엠보 화장솜 1000매 대용량 (500매 2입)', '피부백서 순면 엠보 화장솜 1000매 대용량 (500매 2입)'): '♡4W_PB0013_엠보 소프트 화장솜 2팩 (총 1,000매)',
('살림백서 발각질 제거제 250ml 풋 필링 스프레이 타입 스크럽 연화제', '살림백서 발각질 제거제 250ml 풋 필링 스프레이 타입 스크럽 연화제'): '♡4W_OPP008_발각질 제거제 250ml 1개입',
('1+1 살림백서 뿌리는 곰팡이제거제 400ml', '1+1 살림백서 뿌리는 곰팡이제거제 400ml'): '♥4W_SB0014_뿌리는 곰팡이제거제 400ml 2개입',
('살림백서 제습제 습기제거제 520mlx24개', '살림백서 제습제 습기제거제 520mlx24개'): 'SB_3PL_1/ 살림백서 제습제 1BOX',
('살림백서 발각질 제거제 250ml 풋 필링 스프레이 타입', '살림백서 발각질 제거제 250ml 풋 필링 스프레이 타입'): '♡4W_OPP008_발각질 제거제 250ml 1개입',
('1+1 살림백서 세탁조 클리너 드럼세탁기 청소 통돌이 통세척 (총 8회분)', '1+1 살림백서 세탁조 클리너 드럼세탁기 청소 통돌이 통세척 (총 8회분)'): '♥5M_SB0003_세탁조 클리너 2박스',
('1+1 살림백서 과탄산소다 3kg 대용량 리필형', '1+1 살림백서 과탄산소다 3kg 대용량 리필형'): '■4WC_SBE002_과탄산소다 3kg 2개입',
('살림백서 흰옷 표백제 1kg X 2 EA', '살림백서 흰옷 표백제 1kg X 2 EA'): '살림백서 흰옷 표백제 1kg X 2 EA',
('살림백서 옷걸이형 제습제 옷장형 습기제거제 150g 10개입', '살림백서 옷걸이형 제습제 옷장형 습기제거제 150g 10개입'): '■4MC_SB0084_옷걸이형 제습제 (비드형) 10개입',
('1+1 살림백서 주방세제 청귤향 1L(펌프포함)', '1+1 살림백서 주방세제 청귤향 1L(펌프포함)'): '♥3MC_SB0009_주방세제 1L (청귤향) 2개입',
('1+1 살림백서 산소계 표백제 1kg 흰옷 얼룩제거', '1+1 살림백서 산소계 표백제 1kg 흰옷 얼룩제거'): '■4WC_SBH001_흰옷표백제 1kg 2개입',
('1+1 살림백서 세탁조 클리너 드럼세탁기 청소 통돌이 통세척 8회분 (150g x 8개)', '1+1 살림백서 세탁조 클리너 드럼세탁기 청소 통돌이 통세척 8회분 (150g x 8개)'): '♥5M_SB0003_세탁조 클리너 2박스',
('살림백서 식기세척기 린스 헹굼보조제 가정용 액상형 1L sk lg 밀레 삼성 호환', '살림백서 식기세척기 린스 헹굼보조제 가정용 액상형 1L sk lg 밀레 삼성 호환'): '♥3M_SB0040_식세기 헹굼보조제 1L 1개입',
('1+1 살림백서 과탄산소다 3kg (총 6kg 대용량)', '1+1 살림백서 과탄산소다 3kg (총 6kg 대용량)'): '■4WC_SBE002_과탄산소다 3kg 2개입',
('살림백서 만능 얼룩제거제 리필형 1L 2개입', '살림백서 만능 얼룩제거제 리필형 1L 2개입'): '♥3MC_SB0020_얼룩제거제 500ml 2개입',
('살림백서 제습제 520ml 24개입 1box 옷장 화장실 신발장', '살림백서 제습제 520ml 24개입 1box 옷장 화장실 신발장'): 'SB_3PL_1/ 살림백서 제습제 1BOX',
('1+1 살림백서 탈모 샴푸 엑티브B7 맥주효모 앤 비오틴 1000ml', '1+1 살림백서 탈모 샴푸 엑티브B7 맥주효모 앤 비오틴 1000ml'): '■4WC_OPP004_탈모샴푸 1L 2개입',
('1+1 살림백서 화장실 욕실세정제 800ml 청소세제 곰팡이제거 소독제', '1+1 살림백서 화장실 욕실세정제 800ml 청소세제 곰팡이제거 소독제'): '♥4W_SB0082_욕실세정제 800ml 2개입',
('피부백서 순면 엠보 화장솜 1000매 대용량 (500매 2입)', '피부백서 순면 엠보 화장솜 1000매 대용량 (500매 2입)'): '♡4W_PB0013_엠보 소프트 화장솜 2팩 (총 1,000매)',
('살림백서 발각질 제거제 250ml 풋 필링 스프레이 타입 스크럽 연화제', '살림백서 발각질 제거제 250ml 풋 필링 스프레이 타입 스크럽 연화제'): '♡4W_OPP008_발각질 제거제 250ml 1개입',
('1+1 살림백서 뿌리는 곰팡이제거제 400ml', '1+1 살림백서 뿌리는 곰팡이제거제 400ml'): '♥4W_SB0014_뿌리는 곰팡이제거제 400ml 2개입',







    # 추가적인 조건에 따라 더 많은 값들을 설정할 수 있습니다.
    # , ...
}

highlight_indices = []
# M1과 N1 열의 데이터에 따라 다른 값으로 변경 (첫 번째 행부터 비교 시작)
for index in range(len(df)):
    current_row = df.iloc[index]

    # '상품명'에 해당하는 열 찾기
    m_column = [col for col in df.columns if '상품명' in col or 'product' in col]
    if not m_column:
        raise ValueError(f"No '상품명' or 'product' column found in the dataframe. Columns: {df.columns}")

    # '옵션명'에 해당하는 열 찾기
    n_column = [col for col in df.columns if '옵션명' in col or 'option' in col]
    if not n_column:
        raise ValueError(f"No '옵션명' or 'option' column found in the dataframe. Columns: {df.columns}")

    # 조건에 따라 'O' 열 자동으로 입력
    key = (current_row[m_column[0]], current_row[n_column[0]])
    if key in mapping_dict:
        df.at[index, 'O'] = mapping_dict[key]
    else:
        highlight_indices.append(index)

# 변경된 데이터프레임을 새로운 엑셀 파일로 저장
# df_style = df.style
# df_style.applymap(lambda x: f"background-color: {'lightgrey'}", subset=pd.IndexSlice[highlight_indices, ['N']])
# df_style.to_excel(output_file_path, index=False)

# print("작업이 완료되었습니다.")

# if 'N' in df.columns:
#     df_style = df.style
#     df_style.applymap(lambda x: f"background-color: {'lightgrey'}", subset=pd.IndexSlice[highlight_indices, ['N']])
#     df_style.to_excel(output_file_path, index=False)
#     print("작업이 완료되었습니다.")
# else:
#     raise ValueError("No 'N' column found in the dataframe. Columns: {df.columns}")

# 'N' 열이 존재하는 경우에만 스타일 적용
# if any('N' in col.upper() for col in df.columns):
#     df_style = df.style
#     df_style.applymap(lambda x: f"background-color: {'lightgrey'}", subset=pd.IndexSlice[highlight_indices, [col for col in df.columns if 'N' in col.upper()]])
#     df_style.to_excel(output_file_path, index=False)
#     print("작업이 완료되었습니다.")
# else:
#     raise ValueError("No 'N' column found in the dataframe. Columns: {df.columns}")
#pip install Jinja2


# 'N' 열이 존재하는 경우에만 스타일 적용
if any('N' in col.upper() for col in df.columns):
    df_style = df.style
    df_style = df_style.applymap(lambda x: f"background-color: {'lightgrey'}", subset=pd.IndexSlice[highlight_indices, [col for col in df.columns if 'N' in col.upper()]])
    df_style.to_excel(output_file_path, index=False)
    print("작업이 완료되었습니다.")

    # 메시지 박스 표시
    messagebox.showinfo("완료", "작업이 완료되었습니다!")


else:
    raise ValueError("No 'N' column found in the dataframe. Columns: {df.columns}")
