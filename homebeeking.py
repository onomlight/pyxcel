import pandas as pd
from tkinter import messagebox
# 엑셀 파일 경로
file_path = 'bigtest.xlsx'
output_file_path = 'new__새로운_파일.xlsx'

# 바꿀 단어 설정
# 업데이트시 pyinstaller ex.py 입력
# pyinstaller -F ex.py 후 단일 파일 배포

# 엑셀 파일 불러오기
df = pd.read_excel(file_path)

# 'O'열 초기화
df['O'] = ''

# 조건과 매칭될 값을 정의한 딕셔너리
mapping_dict = {


('베이킹소다 주방세제 750ml 4개입 2종향 설거지 대용량', '쉼표365 주방세제 750ml 4개 라즈베리향'): '#4M_OCOB1014_주방세제 750ml (라즈베리향) 4개입 ♣',
('살림백서 싱크대 물막이 물튀김방지 가림막 물받이 실리콘', '살림백서 싱크대 물막이 물튀김방지 가림막 물받이 실리콘'): '■4WC_SB0123_싱크대 물막이 1개입',
('살림백서 에어프라이어 종이호일 90매 (중형/대형) 기름종이 에어프라이 오븐 원형', '사이즈:중형 90매-1개'): '■1M_SBI002_종이호일 중형 90매',
('살림백서 피톤치드 편백수 스프레이 500ml 3개입 새집증후군 제거방법 베이크아웃 집먼지 퇴지제', '살림백서 피톤치드 편백수 스프레이 500ml 3개입 새집증후군 제거방법 베이크아웃 집먼지 퇴지제'): '■3M_SB0078_피톤치드 편백수 500ml 3개입',
('1+1 살림백서 식기세척기 세제 분말형 1kg 식세기 sk lg 밀레 삼성 호환', '1+1 살림백서 식기세척기 세제 분말형 1kg 식세기 sk lg 밀레 삼성 호환'): '■3MC_SB0039_식세기 세제 1kg (분말) 2개입',
('살림백서 딥퍼퓸 섬유유연제 1L X 3개입 5종향 실내건조 향기좋은 대용량', '향:1.딥퍼퓸 섬유유연제 화이트머스크 1L 3개-1개'): '■4WC_SB0128_딥퍼퓸 섬유유연제 (화이트머스크) 1L 3개입',
('살림백서 실내건조 세탁세제 리필 2.1L 4개입 대용량 액체세제 드럼용 일반용', '사이즈:1)실내건조세탁세제 리필 2.1L 4개_일반용-1개'): '■4MC_SB0161_실내건조 세탁세제 리필형 2.1L 4개입 (일반)',
('살림백서 퍼퓸 섬유유연제 리필형 2L X 3개입', '옵션선택)=2.퍼퓸 섬유유연제 리필 2Lx3개 베베코튼'): '■4WC_SBB003_섬유유연제 2L 리필 (베베코튼) 3개입',
('(단하루!) 살림백서 베이킹소다 3kg 대용량 리필형 1+1', '(단하루!) 살림백서 베이킹소다 3kg 대용량 리필형 1+1'): '■4WC_SBF002_베이킹소다 3kg 2개입',
('살림백서 베이킹소다 3kg 리필형1+1, 3000g, 1개', '1개 3000g'): '■4WC_SBF002_베이킹소다 3kg 2개입',
('베이킹소다 대용량 세탁세제 2.5L 4개입 중성세제 얼룩제거 냄새제거 일반 드럼 세탁기겸용', '베이킹소다 세탁세제 2.5L 4개'): '■4WC_OCOA1003_베이킹소다 세탁세제 2.5L 4개입 ♣',
('살림백서 딥클린 식기세척기 린스 1L 헹굼보조제 전용세제', '옵션 선택=딥클린 식기세척기 린스 1L'): '★4M_SBDC005_딥클린 식세기 헹굼보조제 1L 1개입',
('살림백서 베이비&키즈 아기로션 500ml 유아로션 호호바 온가족 어린이 키즈 패밀리', '살림백서 베이비&키즈 아기로션 500ml 유아로션 호호바 온가족 어린이 키즈 패밀리'): '♡4W_BA0002_아기 로션 500ml 1개입',
('살림백서 발각질 제거제 250ml 풋 필링 스프레이 타입 스크럽 연화제', '살림백서 발각질 제거제 250ml 풋 필링 스프레이 타입 스크럽 연화제'): '♡4W_OPP008_발각질 제거제 250ml 1개입',
('1+1 살림백서 오푼티아&밤부 샴푸 500ml', '향 선택=01.베이비파우더향'): '♡4W_OPS0001_샴푸 (베이비파우더) 500ml 2개입',
('1+1 살림백서 천연 지수81% 샴푸 500ml 약산성 퍼퓸 비듬 향기좋은 사춘기 청소년 초등학생 지성 두피 정수리냄새 대용량 오푼티아 밤부', '향:화이트머스크향 샴푸 500ml 1+1-1개'): '♡4W_OPS0002_샴푸 (화이트머스크) 500ml 2개입',
('살림백서 대용량 바디로션 라이스&허브 1000ml 향좋은 고보습 촉촉한 끈적임없는 냄새좋은', '향:RH1.딥그린시더우드 바디로션 1개-1개'): '♡4W_RH_OPR013_R허브 바디로션 딥그린시더우드 1L 1개입 ',
('1+1 살림백서 대용량 바디로션 라이스&허브 1000ml 향좋은 고보습 촉촉한 끈적임없는', '향:RH1.딥그린시더우드 바디로션 2개-1개'): '♡4W_RH_OPR014_R허브 바디로션 딥그린시더우드 1L 2개입',
('살림백서 쌀겨 대용량 클렌징폼 500ml 세안제 저자극 딥클렌징 지성 민감성 임산부', '향:1)살림백서 쌀겨 클렌징폼 500ml-1개'): '♡4W_RH_OPR029_R허브 쌀겨 클렌징폼 500ml 1개입',
('살림백서 토탄수 대용량 클렌징폼 500ml 세안제 저자극 딥클렌징 지성 민감성 임산부', '향:1)살림백서 토탄수 클렌징폼 500ml-1개'): '♡4W_RH_OPR031_R허브 토탄수 클렌징폼 500ml 1개입',
('1+1 살림백서 차량용 방향제 디퓨저 80ml 6종향 고급 자동차 차량 명품', '구성품:1+1 살림백서 차량방향제 블랙체리-1개'): '♡4W_RH_SB0148_차량방향제 (블랙체리) 80ml 2개입',
('(단7일! 행사가) 마미바티 아기세제 1+1', '(단7일! 행사가) 마미바티 아기세제 1+1'): '♥3M_OMVA1002_마미바티 아기세제 2개입',
('천연유래 아기세제 1.8L 2개', '마미바티 아기세제 2개입'): '♥3M_OMVA1002_마미바티 아기세제 2개입',
('1+1 마미바티 아기세제 1.8L 유아 세탁세제 신생아', '선택1: 01_1+1 마미바티 아기세제1800ml'): '♥3M_OMVA1002_마미바티 아기세제 2개입',
('마미바티 아기 섬유유연제 1+1 (단7일 행사가)', '마미바티 아기 섬유유연제 1+1 (단7일 행사가)'): '♥3M_OMVA1005_마미바티 섬유유연제 2개입',
('아기 섬유유연제 1.8L 2개', '마미바티 아기섬유유연제 1.8L 2개'): '♥3M_OMVA1005_마미바티 섬유유연제 2개입',
('1+1 마미바티 아기세제 1.8L 유아 세탁세제 신생아', '선택1: 02_1+1 마미바티 아기유연제1800ml'): '♥3M_OMVA1005_마미바티 섬유유연제 2개입',
('살림백서 1+1 식기세척기 세제 액상형 1L', '제품 선택=식기세척기 세제 액상형 1L 2개'): '♥3M_SB0037_식세기 세제 1L (액체) 2개입',
('살림백서 하루세번 구취케어 치약 100g 3개입 입냄새 구취제거 잇몸', '사이즈:02.하루세번 잇몸케어 치약 100g 3개입-1개'): '♥3M_SB0155_잇몸케어 치약 3개입',
('(단7일! 행사가) 마미바티 아기 젖병세정제 1 + 1  500ml', '(단7일! 행사가) 마미바티 아기 젖병세정제 1 + 1  500ml'): '♥4MC_OMVB1002_마미바티 주방세제 2개입',
('살림백서 캡슐세제 고농축 트리플 3in1 30개입 액체 세탁세제 드럼일반겸용', '살림백서 캡슐세제 고농축 트리플 3in1 30개입 액체 세탁세제 드럼일반겸용'): '♥4MC_SB0063_캡슐세제 30개입 1팩',
('살림백서 소프트 퍼퓸 건조기시트 80매 건조기섬유유연제 드라이시트 3종향', '향:3)퍼퓸 건조기 섬유유연제 80매(코트니향)-1개'): '♥4MC_SB0076_건조기 섬유유연제 (코트니향) 2박스',
('살림백서 2단계 초 미세모 칫솔 딥그래이쉬 20개입 초극세모 임산부 대용량', '살림백서 2단계 초 미세모 칫솔 딥그래이쉬 20개입 초극세모 임산부 대용량'): '♥4MC_SB0126_딥그레이쉬 칫솔 20개입 1박스',
('살림백서 세탁세제 1+1', '살림백서 세탁세제 1+1'): '♥4WC_SBA001_세탁세제 2L (본품) 2개입',
('살림백서 배수구클리너  600g 2팩 (총 150g 8개입 / 8회분)', '살림백서 배수구클리너  600g 2팩 (총 150g 8개입 / 8회분)'): '♥4M_SB0001_배수구 클리너 2박스',
('1+1 살림백서 세탁조 클리너 드럼세탁기 청소 통돌이 통세척 (총 8회분)', '1+1 살림백서 세탁조 클리너 드럼세탁기 청소 통돌이 통세척 (총 8회분)'): '♥4M_SB0003_세탁조 클리너 2박스',
('뽀드득 유리세정제 750ml 유리 거울 물때 얼룩제거', '유리세정제 750ml 1개'): '#4M_OCOB1023_유리세정제 750ml 1개입 ♣',
('깔끔말끔 뿌리는 곰팡이제거제 750ml 베란다 욕실 화장실 벽 창틀 청소세제', '뿌리는 곰팡이제거제 750ml 1개'): '#4M_OCOB1026_곰팡이제거제 750ml 1개입 ♣',
('살림백서 라이스&허브 1종 주방세제 750ml 3개입', '제품 선택=라이스&허브 주방세제 750ml 3개입, 향기 선택=스위트허브향'): '■○3M_SB0109_R허브 주방세제 (스위트허브) 750ml 3개입',
('마미바티 아기 섬유유연제 1800ml 3개입 유아 신생아', '옵션 선택): 1)마미바티 아기섬유유연제 3개입'): '■♥3M_OMVA1006_마미바티 섬유유연제 3개입',
('살림백서 에어프라이어 종이호일 90매 (중형/대형) 기름종이 에어프라이 오븐 원형', '사이즈:대형 90매-1개 (+1000원)'): '■4WC_SBI001_종이호일 대형 90매 1박스',
('살림백서 에어프라이어 종이호일 90매 (중형/대형) 기름종이 에어프라이 오븐 원형', '사이즈:대형 90매-2개 (+2000원)'): '■4WC_SBI001_종이호일 대형 90매 1박스',
('살림백서 고농축 섬유유연제 1LX3개입 퍼퓸 향기좋은 대용량 향좋은 실내건조', '향:8)스위트베리+엠버화이트+아쿠아프레쉬 각1개-1개'): '■3MC_SB0092_스위트베리+엠버화이트+아쿠아프레쉬 각 1개입',
('살림백서 딥퍼퓸 섬유유연제 1L x 3개입 5종향', '제품 선택=딥퍼퓸 섬유유연제 1L 3개, 향기 선택=베이비파우더'): '■4WC_SB0128_딥퍼퓸 섬유유연제 (화이트머스크) 1L 3개입',
('살림백서 옷걸이형 제습제 옷장 옷걸이 습기제거제 150g 10개입', '옵션 선택=옷걸이형  습기제거제 150g 10개입'): '■3M_SB0084_옷걸이형 제습제 (비드형) 10개입',
('살림백서 대용량 1종 주방세제 라이스&허브 4L+750ml', '살림백서 라이스&허브 주방세제 4L+750ml=1)주방세제 석류향 4L+750ml'): '■4WC_SB0106_R허브 주방세제 (석류) 4L 1개입+750ml 1개입',
('1+1 살림백서 엑티브B7 맥주효모 앤 비오틴 탈모완화 샴푸 1000ml 대용량 지성 두피 남자 여자 약산성', '1+1 살림백서 엑티브B7 맥주효모 앤 비오틴 탈모완화 샴푸 1000ml 대용량 지성 두피 남자 여자 약산성'): '■4WC_OPP004_탈모샴푸 1L 2개입',
('살림백서 주방세제 리필형 1L 3개입 대용량 천연유래98% 리필 과일세척', '향:살림백서 주방세제 리필 1Lx3개(청귤)-1개'): '■4WC_SB0010_주방세제 리필형 1L (청귤향) 3개입',
('살림백서 실내건조 세탁세제 2.7L 2개입(일반용) 대용량 액체세제', '살림백서 실내건조 세탁세제 2.7L 2개입(일반용) 대용량 액체세제'): '■4WC_SB0099_실내건조 세탁세제 (일반) 2.7L 2개입',
('살림백서 실내건조 세탁세제 2.7L x 4개입 대용량 액체세제 드럼용 일반용', '상품선택=1)실내건조 세탁세제 2.7L 4개입(일반용)'): '■4WC_SB0100_실내건조 세탁세제 (일반) 2.7L 4개입',
('살림백서 실내건조 세탁세제 2.7L x 4개입 대용량 액체세제 드럼용 일반용', '상품선택=4)실내건조 세탁세제 2.7L 2개입(드럼용)'): '■4WC_SB0101_실내건조 세탁세제 (드럼) 2.7L 2개입',
('살림백서 퍼퓸 섬유유연제 리필형 2L X 3개입', '옵션선택)=3.섬유유연제 리필 오리지널 2Lx3개 한라봉'): '■4WC_SBB004_섬유유연제 2L 리필 (한라봉) 3개입',
('1+1 살림백서 과탄산소다 3kg 대용량 리필형', '1+1 살림백서 과탄산소다 3kg 대용량 리필형'): '■4WC_SBE002_과탄산소다 3kg 2개입',
('살림백서 구연산 1kg 리필형 1+1', '살림백서 구연산 1kg 리필형 1+1'): '■4WC_SBG002_구연산 1kg 2개입',
('1+1 살림백서 구연산 1kg', '1+1 살림백서 구연산 1kg'): '■4WC_SBG002_구연산 1kg 2개입',
('1+1 살림백서 딥클린 곰팡이 제거제 800ml 대용량 벽 화장실 욕실 실리콘 창틀 청소세제', '1+1 살림백서 딥클린 곰팡이 제거제 800ml 대용량 벽 화장실 욕실 실리콘 창틀 청소세제'): '★4M_SBDC002_딥클린 곰팡이제거제 800ml 2개입',
('살림백서 1+1 딥클린 곰팡이 제거제 800ml 대용량', '옵션 선택=딥클린 곰팡이 제거제 800ml 2개'): '★4M_SBDC002_딥클린 곰팡이제거제 800ml 2개입',
('살림백서 1+1 딥클린 욕실세정제 800ml', '옵션 선택=딥클린 욕실세정제 800ml 2개'): '★4M_SBDC003_딥클린 욕실세정제 800ml 2개입',
('살림백서 딥클린 1종 주방세제 750ml 3개입', '제품 선택=딥클린 주방세제 750ml 2개, 향기 선택=프레쉬베리'): '★4M_SBDC009_딥클린 주방세제 그린허브 750ml 2개입',
('살림백서 베이비&키즈 올인원 바디워시 500ml 아기 유아 신생아 헤어 바디 바스앤샴푸', '살림백서 베이비&키즈 올인원 바디워시 500ml 아기 유아 신생아 헤어 바디 바스앤샴푸'): '♡4W_BA0003_아기 바디워시 500ml 1개입',
('1+1 살림백서 오푼티아 바디로션 500ml 대용량 저자극 끈적임없는 천연 지수 72% 향기좋은 향수 퍼퓸', '향:03.1+1 바디로션 베베머스크 500ml-1개'): '♡4W_OPBL004_바디로션 (베베머스크향) 500ml 2개입',
('1+1 살림백서 오푼티아 바디로션 500ml 대용량 저자극 끈적임없는 천연 지수 72% 향기좋은 향수 퍼퓸', '향:01.1+1 바디로션 무향 500ml-1개'): '♡4W_OPBL008_바디로션 (무향) 500ml 2개입',
('1+1 살림백서 오푼티아&밤부 바디로션 500ml', '바디로션 향 선택1=02.바디로션 화이트머스크, 바디로션 향 선택2=03.바디로션 베베머스크'): '♡4W_OPBL105_바디로션 화이트머스크+베베머스크 500ml',
('살림백서 라이스허브 트리트먼트 딥그린시더우드 1L 2개', '살림백서 라이스허브 트리트먼트 딥그린시더우드 1L 2개'): '♡4W_RH_OPR006_R허브 트리 딥그린시더우드 1L 2개입',
('1+1 살림백서 라이스 앤 허브 바디워시 1L 대용량 퍼퓸 향좋은 샤워젤 바디클렌져', '향:RH1.딥그린시더우드 바디워시 (1+1)-1개'): '♡4W_RH_OPR010_R허브 바디워시 딥그린시더우드 1L 2개입',
('살림백서 토탄수 대용량 클렌징폼 500ml 세안제 저자극 딥클렌징 지성 민감성 임산부', '향:2)살림백서 쌀겨 클렌징폼 500ml-1개'): '♡4W_RH_OPR029_R허브 쌀겨 클렌징폼 500ml 1개입',
('1+1 살림백서 디퓨저 라이스&허브 200ml 아로마 실내 방향제 화장실 인테리어 대용량', '향:4)1+1 살림백서 디퓨저 포레스트가든-1개'): '♡4W_RH_SB0115_R허브 디퓨저 (포레스트가든) 200ml 2개입',
('1+1 살림백서 디퓨저 라이스&허브 200ml 아로마 실내 방향제 화장실 인테리어 대용량', '향:1)1+1 살림백서 디퓨저 일랑일랑-2개'): '♡4W_RH_SB0118_R허브 디퓨저 (일랑일랑) 200ml 2개입',
('1+1 살림백서 라이스&허브 디퓨저 200ml (아로마 실내 방향제)', '향 선택=5)살림백서 디퓨저 1+1 베이비파우더'): '♡4W_RH_SB0120_R허브 디퓨저 (화이트머스크) 200ml 2개입',
('1+1 살림백서 디퓨저 라이스&허브 200ml 아로마 실내 방향제 화장실 인테리어 대용량', '향:5)1+1 살림백서 디퓨저 베이비파우더-1개'): '♡4W_RH_SB0120_R허브 디퓨저 (화이트머스크) 200ml 2개입',
('1+1 살림백서 차량용 방향제 디퓨저 80ml 6종향 고급 자동차 차량 명품', '구성품:1+1 살림백서 차량방향제 포레스트가든-1개'): '♡4W_RH_SB0149_차량방향제 (포레스트가든) 80ml 2개입',
('살림백서 차량용 방향제 디퓨저 80ml 6종향', '제품선택=차량용 디퓨저 80ml 2개입, 향기선택=베이비파우더'): '♡4W_RH_SB0152_차량방향제 (베이비파우더) 80ml 2개입',
('1+1 살림백서 라이스 앤 모링가 바디워시 1500ml 대용량 냄새좋은 향좋은 샤워젤 바디클렌져', '향:RM2.베이비파우더향 바디워시 1500ml (1+1)-1개'): '♡4W_RM_OPM008_R모링가 바디워시 베이비파우더향 1.5L 2개입',
('1+1 살림백서 베이비&키즈 아기젖병세정제 무향 500ml 아기 유아 식기 세제', '1+1 살림백서 베이비&키즈 아기젖병세정제 무향 500ml 아기 유아 식기 세제'): '♥3M_BA0012_베이비 포밍 주방세제  500ml 2개입',
('마미바티 아기세제 & 섬유유연제 세트', '마미바티 아기세제 & 섬유유연제 세트'): '♥3M_OMVA1007_마미바티 세제1개+유연제1개 SET',
('살림백서 과탄산소다 (용기형) 400g', '살림백서 과탄산소다 (용기형) 400g'): '♥3M_SB0055_과탄산소다 400g 용기형 1개입',
('살림백서 구연산 (용기형) 400g', '살림백서 구연산 (용기형) 400g'): '♥3M_SB0057_구연산 400g 용기형 1개입',
('살림백서 베이킹소다 (용기형) 450g', '살림백서 베이킹소다 (용기형) 450g'): '♥3M_SB0059_베이킹소다 450g 용기형 1개입',
('살림백서 피톤치드 편백수 스프레이 500ml 3개입 새집증후군 제거방법 베이크아웃 집먼지 퇴지제', '살림백서 피톤치드 편백수 스프레이 500ml 3개입 새집증후군 제거방법 베이크아웃 집먼지 퇴지제'): '■3M_SB0078_피톤치드 편백수 500ml 3개입',
('살림백서 액티브 식기세척기 클리너 100g 4회분 식세기 청소 세제 세정제', '살림백서 액티브 식기세척기 클리너 100g 4회분 식세기 청소 세제 세정제'): '♥3M_SB0096_식기세척기 클리너 100g 4개입 1박스',
('살림백서 하루세번 구취케어 치약 100g 3개입 입냄새 구취제거 잇몸', '사이즈:01.하루세번 구취케어 치약 100g 3개입-1개'): '♥3M_SB0154_구취케어 치약 3개입',
('주방세제 소분 용기 (350m)', '주방세제 소분 용기 (350m)'): '♥3M_SBAD01_주방세제 용기 1개입 * ',
('1+1 살림백서 주방세제 1L', '향 선택=자몽향'): '♥4WC_SB0006_주방세제 1L (자몽향) 2개입',
('살림백서 주방세제 청귤향 1L 2개입(펌프포함)', '살림백서 주방세제 청귤향 1L 2개입(펌프포함)'): '♥4WC_SB0009_주방세제 1L (청귤향) 2개입',
('1+1 살림백서 주방세제 1L', '향 선택=청귤향'): '♥4WC_SB0009_주방세제 1L (청귤향) 2개입',
('살림백서 주방세제 1L 1+1', '살림백서 주방세제 1L 1+1'): '♥4WC_SB0009_주방세제 1L (청귤향) 2개입',
('1+1 살림백서 울세제 1L', '1+1 살림백서 울세제 1L'): '♥4WC_SB0012_울세제 1L 2개입',
('1+1 살림백서 울세제 중성세제 1L 홈드라이 세탁', '1+1 살림백서 울세제 중성세제 1L 홈드라이 세탁'): '♥4WC_SB0012_울세제 1L 2개입',
('1+1 살림백서 기름때 클리너 400ml', '1+1 살림백서 기름때 클리너 400ml'): '♥4M_SB0018_기름때 클리너 400ml 2개입',
('1+1 살림백서 기름때 클리너 400ml 가스렌지 후드 청소 주방 다목적 클리너', '1+1 살림백서 기름때 클리너 400ml 가스렌지 후드 청소 주방 다목적 클리너'): '♥4M_SB0018_기름때 클리너 400ml 2개입',
('1+1 살림백서 만능 옷 얼룩제거제 흰옷얼룩제거 500ml', '1+1 살림백서 만능 옷 얼룩제거제 흰옷얼룩제거 500ml'): '♥3MC_SB0020_얼룩제거제 500ml 2개입',
('1+1 살림백서 식기세척기 세제 분말형 1kg 식세기 sk lg 밀레 삼성 호환', '1+1 살림백서 식기세척기 세제 분말형 1kg 식세기 sk lg 밀레 삼성 호환'): '■3MC_SB0039_식세기 세제 1kg (분말) 2개입',
('살림백서 섬유탈취제 500ml 4개입 섬유향수 드레스퍼퓸 스프레이 탈취99% 정전기방지', '향:살림백서 섬유탈취제 500ml 4개(블랑)-1개'): '♥3MC_SB0080_섬유탈취제 (코지블랑) 500ml 2개입',
('살림백서 딥퍼퓸 섬유유연제 1L x 3개입 5종향', '제품 선택=딥퍼퓸 섬유유연제 1L 3개, 향기 선택=화이트머스크'): '■4WC_SB0128_딥퍼퓸 섬유유연제 (화이트머스크) 1L 3개입',
('1+1 살림백서 딥퍼퓸 섬유탈취제 500ml 7종 섬유향수 정전기방지 스프레이 드레스', '향:1)1+1 화이트머스크 500ml-1개'): '■♥4WC_SB0139_딥퍼퓸 섬유탈취제 (화이트머스크) 500ml 2개입',
('2+1 마미바티 아기 젖병세정제 500g 1종 주방세제 유아식기', '옵션선택 (택1): 01.2+1 마미바티 아기 주방세제 500g'): '♥4MC_OMVB1003_마미바티 주방세제 3개입',
('살림백서 찌든때 클리너 240g 2개입 스테인레스 연마제거제 탄냄비닦는법 텀블러 세척', '살림백서 찌든때 클리너 240g 2개입 스테인레스 연마제거제 탄냄비닦는법 텀블러 세척'): '♥3M_SB0019_찌든때 클리너 1박스',
('살림백서 변기세정제 (40g x 10개)', '살림백서 변기세정제 (40g x 10개)'): '♥4MC_SB0050_변기세정제 1박스',
('살림백서 소프트 퍼퓸 건조기 섬유유연제 40매 3종향', '향 선택)=4)퍼퓸 건조기 섬유유연제 80매(블랑향)'): '♥4MC_SB0072_건조기 섬유유연제 (블랑향) 2박스',
('살림백서 소프트 퍼퓸 건조기 섬유유연제 40매 3종향', '향 선택)=6)퍼퓸 건조기 섬유유연제 80매(코트니)'): '♥4MC_SB0076_건조기 섬유유연제 (코트니향) 2박스',
('1+1 살림백서 액체세탁세제 2L', '1+1 살림백서 액체세탁세제 2L'): '♥4WC_SBA001_세탁세제 2L (본품) 2개입',
('1+1 살림백서 천연유래98% 세탁세제', '1+1 살림백서 천연유래98% 세탁세제'): '♥4WC_SBA001_세탁세제 2L (본품) 2개입',
('(단하루!) 살림백서1+1 세탁세제 섬유유연제 주방세제 베이킹', '옵션명 1:02_살림백서 세탁세제 1+1-1개'): '♥4WC_SBA001_세탁세제 2L (본품) 2개입',
('1+1 살림백서 섬유유연제 2L', '1+1 살림백서 섬유유연제 2L'): '♥4MC_SBB001_섬유유연제 2L (본품) 2개입',
('1+1 살림백서 섬유유연제 2L 2개', '1+1 살림백서 섬유유연제 2L 2개'): '♥4MC_SBB001_섬유유연제 2L (본품) 2개입',
('(단하루!) 살림백서1+1 세탁세제 섬유유연제 주방세제 베이킹', '옵션명 1:03_살림백서 섬유유연제 1+1-1개'): '♥4MC_SBB001_섬유유연제 2L (본품) 2개입',
('1+1 살림백서 뿌리는 곰팡이제거제 400ml', '1+1 살림백서 뿌리는 곰팡이제거제 400ml'): '♥4WC_SB0014_뿌리는 곰팡이제거제 400ml 2개입',
('1+1 살림백서 곰팡이제거제 젤 벽 벽지 화장실 베란다 실리콘 창틀 욕실 결로 곰팡이', '옵션선택):1+1 살림백서 뿌리는 곰팡이제거제-1개'): '♥4WC_SB0014_뿌리는 곰팡이제거제 400ml 2개입',
('1+1 살림백서 화장실 욕실세정제 800ml 청소세제 곰팡이제거 소독제', '1+1 살림백서 화장실 욕실세정제 800ml 청소세제 곰팡이제거 소독제'): '■4WC_SB0082_욕실세정제 800ml 2개입',
('1+1 살림백서 배수구 클리너 (총 8회분)', '1+1 살림백서 배수구 클리너 (총 8회분)'): '♥4M_SB0001_배수구 클리너 2박스',
('1+1 살림백서 배수구 클리너 150g 총 8개입 싱크대청소 하수구 세척', '1+1 살림백서 배수구 클리너 150g 총 8개입 싱크대청소 하수구 세척'): '♥4M_SB0001_배수구 클리너 2박스',
('1+1 살림백서 세탁조 클리너 (총 8회분)', '1+1 살림백서 세탁조 클리너 (총 8회분)'): '♥4M_SB0003_세탁조 클리너 2박스',
('1+1 살림백서 차량용 방향제 디퓨저 80ml 6종향 고급 자동차 차량 명품', '구성품:1+1 살림백서 차량방향제 화이트머스크-1개'): '♡4W_RH_SB0153_차량방향제 (화이트머스크) 80ml 2개입',
('살림백서 대용량 1종 주방세제 라이스&허브 4L+750ml', '살림백서 라이스&허브 주방세제 4L+750ml=2)주방세제 스위트허브향 4L+750ml'): '■4W_SB0111_R허브 주방세제 (스위트허브) 4L 1개입+750ml 1개입',
('1+1 살림백서 라이스&허브 버블 핸드워시 500m', '제품 선택)=01.라이스&허브 핸드워시 500ml 2개, 향기 선택)=RH1.딥그린시더우드향'): '♡4W_RH_OPR018_R허브 핸드워시 딥그린시더우드 500ml 2개입',
('1+1 살림백서 라이스&허브 버블 핸드워시 500m', '제품 선택)=01.라이스&허브 핸드워시 500ml 2개, 향기 선택)=RH2.로지피치향'): '♡4W_RH_OPR021_R허브 핸드워시 로지피치 500ml 2개입',
('1+1 살림백서 라이스&허브 버블 핸드워시 500m', '제품 선택)=01.라이스&허브 핸드워시 500ml 2개, 향기 선택)=RH3.그린포레스트향'): '♡4W_RH_OPR024_R허브 핸드워시 그린포레스트 500ml 2개입',
('1+1 살림백서 라이스&허브 버블 핸드워시 500m', '제품 선택)=01.라이스&허브 핸드워시 500ml 2개, 향기 선택)=RH4.청포도향'): '♡4W_RH_OPR027_R허브 핸드워시 청포도 500ml 2개입',
('살림백서 차량용 방향제 디퓨저 80ml 6종향', '제품선택=차량용 디퓨저 80ml 2개입, 향기선택=포레스트가든'): '♡4W_RH_SB0149_차량방향제 (포레스트가든) 80ml 2개입',
('1+1 살림백서 디퓨저 라이스&허브 200ml 아로마 실내 방향제 화장실 인테리어 대용량', '향:1)1+1 살림백서 디퓨저 일랑일랑-1개'): '♡4W_RH_SB0118_R허브 디퓨저 (일랑일랑) 200ml 2개입',
('살림백서 소프트 퍼퓸 건조기시트 80매 건조기섬유유연제 드라이시트 3종향', '향:2)퍼퓸 건조기 섬유유연제80매(러브미향)-1개'): '♥3M_SB0074_건조기 섬유유연제 (러브미향) 2박스',
('1+1 살림백서 디퓨저 라이스&허브 200ml 아로마 실내 방향제 화장실 인테리어 대용량', '향:3)1+1 살림백서 디퓨저 블랙체리-1개'): '♡4W_RH_SB0117_R허브 디퓨저 (블랙체리) 200ml 2개입',
('촉촉 비데티슈 비데물티슈 60매 10팩 화장실 화장실용 비데용 물에녹는', '쉼표365 비데물티슈 60매 X 10팩'): '■#1M_OCOA2001_쉼표 비데 물티슈 60매 10개입 ♣',
('살림백서 딥퍼퓸 섬유유연제 1L x 3개입 5종향', '제품 선택=딥퍼퓸 섬유유연제 1L 3개, 향기 선택=스위트일랑'): '■3MC_SB0130_딥퍼퓸 섬유유연제 (소프트코튼) 1L 3개입',
('1+1 살림백서 오푼티아 바디로션 500ml 대용량 저자극 끈적임없는 천연 지수 72% 향기좋은 향수 퍼퓸', '향:02.1+1 바디로션 화이트머스크 500ml-1개'): '♡4W_OPBL002_바디로션 (화이트머스크향) 500ml 2개입',
('살림백서 발각질 제거제 250ml 풋 필링 스프레이 타입', '살림백서 발각질 제거제 250ml 풋 필링 스프레이 타입'): '♡4W_OPP008_발각질 제거제 250ml 1개입',
('1+1 살림백서 오푼티아&밤부 샴푸 500ml', '향 선택=05.체리블라썸향'): '♡4W_OPS0005_샴푸 (체리블라썸) 500ml 2개입',
('1+1 살림백서 오푼티아&밤부 샴푸 500ml', '향 선택=06.유칼립투스향'): '♡4W_OPS0006_샴푸 (유칼립투스) 500ml 2개입',
('1+1 살림백서 오푼티아&밤부 트리트먼트 500ml', '향 선택=04.유칼립투스향'): '♡4W_OPT0006_트리 (유칼립투스) 500ml 2개입',
('살림백서 라이스 앤 모링가 바디워시 1500ml 대용량 냄새좋은 향좋은 샤워젤 바디클렌져', '향:RM2.베이비파우더향 바디워시 1500ml 1개-1개'): '♡4W_RM_OPM007_R모링가 바디워시 베이비파우더향 1.5L 1개입',
('1+1 살림백서 라이스 앤 모링가 바디워시 1500ml 대용량 냄새좋은 향좋은 샤워젤 바디클렌져', '향:RM2.베이비파우더향 바디워시 1500ml (1+1)-2개'): '♡4W_RM_OPM008_R모링가 바디워시 베이비파우더향 1.5L 2개입',
('살림백서 제습백서 젤리 제습제 모음전 일반형 옷장걸이형', '구성품:1) 젤리 제습백서 걸이형 1개입 4팩(4개입)-1개'): '♥3M_SB0067_젤리 제습백서 걸이형 (1BOX)',
('살림백서 제습백서 젤리 제습제 모음전 일반형 옷장걸이형', '구성품:1) 젤리 제습백서 걸이형 1개입 4팩(4개입)-5개'): '♥3M_SB0067_젤리 제습백서 걸이형 (1BOX)',
('살림백서 제습백서 젤리 제습제 모음전 일반형 옷장걸이형', '구성품:2) 젤리 제습백서 틈새형 4개입 4팩(16개입)-4개'): '♥3M_SB0069_젤리 제습백서 일반형 (1BOX)',
('살림백서 액티브 식기세척기 클리너 100g 4회분', '살림백서 액티브 식기세척기 클리너 100g 4회분'): '♥3M_SB0096_식기세척기 클리너 100g 4개입 1박스',
('살림백서 하루세번 구취케어 치약 100g 3개입 플로랄민트향', '살림백서 하루세번 구취케어 치약 100g 3개입 플로랄민트향'): '♥3M_SB0154_구취케어 치약 3개입',
('1+1 살림백서 오푼티아&밤부 핸드크림 로션 타입 300ml 퍼퓸 대용량 고보습 향기좋은', '향 선택:화이트머스크향-1개'): '♡4W_OPH001_핸드크림 (화이트머스크향) 300ml 2개입',
('살림백서 찌든때 클리너 240g x 2개입', '살림백서 찌든때 클리너 240g x 2개입'): '♥3M_SB0019_찌든때 클리너 1박스',
('1+1 살림백서 과탄산소다 3kg (총 6kg 대용량)', '1+1 살림백서 과탄산소다 3kg (총 6kg 대용량)'): '■4WC_SBE002_과탄산소다 3kg 2개입',
('살림백서 대용량 바디로션 라이스&허브 1000ml 향좋은 고보습 촉촉한 끈적임없는 냄새좋은', '향:RH2.그린포레스트 바디로션 1개-1개'): '♡4W_RH_OPR015_R허브 바디로션 그린포레스트 1L 1개입',
('1+1 살림백서 디퓨저 라이스&허브 200ml 아로마 실내 방향제 화장실 인테리어 대용량', '향:6)1+1 살림백서 디퓨저 화이트머스크-1개'): '♡4W_RH_SB0120_R허브 디퓨저 (화이트머스크) 200ml 2개입',
('1+1 살림백서 오푼티아&밤부 바디워시 500ml', '향 선택=01.베이비파우더향'): '♡4W_OPB0001_바디워시 (베이비파우더) 500ml 2개입',
('1+1 살림백서 오푼티아 바디워시 500ml 약산성 천연 지수 81% 향기좋은 남자 여자 퍼퓸', '향:유칼립투스향 바디워시 500ml 1+1-1개'): '♡4W_OPB0006_바디워시 (유칼립투스) 500ml 2개입',
('살림백서 과탄산소다 3kg', '살림백서 과탄산소다 3kg'): '■1M_SBE001_과탄산소다 3kg 1개입',
('살림백서 차량용 방향제 디퓨저 80ml 6종향', '제품선택=차량용 디퓨저 80ml 2개입, 향기선택=블랙체리'): '♡4W_RH_SB0148_차량방향제 (블랙체리) 80ml 2개입',
('1+1 살림백서 오푼티아&밤부 바디로션 500ml', '바디로션 향 선택1=04.바디로션 체리블라썸, 바디로션 향 선택2=04.바디로션 체리블라썸'): '♡4W_OPBL006_바디로션 체리블라썸 500ml 2개입 ',
('1+1 살림백서 라이스&허브 디퓨저 200ml (아로마 실내 방향제)', '향 선택=1)살림백서 디퓨저 1+1 포레스트가든'): '♡4W_RH_SB0115_R허브 디퓨저 (포레스트가든) 200ml 2개입',
('마미바티 젖병세정제 500g', '마미바티 젖병세정제 500g'): '♥3M_OMVB1001_마미바티 주방세제 1개입',
('살림백서 소프트 모이스처 세수 비누 100g 10개입', '살림백서 소프트 모이스처 세수 비누 100g 10개입'): '♥4MC_SB0156_세수 비누 100g 10개입',
('살림백서 바르는 곰팡이제거젤 3개입', '살림백서 바르는 곰팡이제거젤 3개입'): '♥4M_SB0016_바르는 곰팡이제거제 150g 3개입',
('1+1 살림백서 차량용 방향제 디퓨저 80ml 6종향 고급 자동차 차량 명품', '구성품:1+1 살림백서 차량방향제 일랑일랑-1개'): '♡4W_RH_SB0151_차량방향제 (일랑일랑) 80ml 2개입',
('1+1 살림백서 다목적세정제 400ml', '1+1 살림백서 다목적세정제 400ml'): '♥3MC_SB0023_다목적 클리너 500ml 2개입',
('1+1 살림백서 오푼티아 바디워시 500ml 약산성 천연 지수 81% 향기좋은 남자 여자 퍼퓸', '향:일랑일랑 바디워시 500ml 1+1-2개'): '♡4W_OPB0004_바디워시 (일랑일랑) 500ml 2개입',
('(단7일! 행사가) 마미바티 아기 젖병세정제 1 + 1  500ml','(단7일! 행사가) 마미바티 아기 젖병세정제 1 + 1  500ml'):'♥4MC_OMVB1002_마미바티 주방세제 2개입',
('살림백서 보들보들 미용티슈 2겹 250매 6입 각티슈  곽휴지 화장지','살림백서 보들보들 미용티슈 2겹 250매 6입 각티슈  곽휴지 화장지'):'SB_3PL_3/ 살림백서 미용티슈 1BOX ',
('살림백서 딥퍼퓸 섬유유연제 1L X 3개입 5종향 실내건조 향기좋은 대용량', '향:4.딥퍼퓸 섬유유연제 스위트일랑 1L 3개-1개'): '■4WC_SB0131_딥퍼퓸 섬유유연제 (스위트일랑) 1L 3개입',
('살림백서 딥퍼퓸 섬유유연제 1L X 3개입 5종향 실내건조 향기좋은 대용량', '향:3.딥퍼퓸 섬유유연제 소프트코튼 1L 3개-1개'): '■4WC_SB0130_딥퍼퓸 섬유유연제 (소프트코튼) 1L 3개입',
('살림백서 딥클린 주방세제 750ml 3개입 대용량 주방세제 설거지', '향:2)딥클린 주방세제 프레쉬베리 750ml 3개-1개'): '★4M_SBDC008_딥클린 주방세제 프레쉬베리 750ml 3개입',
('살림백서 딥클린 주방세제 750ml 3개입 대용량 주방세제 설거지', '향:1)딥클린 주방세제 그린허브 750ml 3개-1개'): '★4M_SBDC007_딥클린 주방세제 그린허브 750ml 3개입',
('쉼표365 식기세척기 린스 헹굼보조제 삼성 LG 엘지 SK 밀레 식세기 호환 식세기 전용', '옵션선택): 01) 식기세척기 린스 500ml'): '#3M_OCOB1004_식세기 린스 500ml 1개입 ♣',
('살림백서 뽑아쓰는 정전기 청소포 200매 일회용 바닥 청소', '살림백서 뽑아쓰는 정전기 청소포 200매 일회용 바닥 청소'): '■4MC_SB0133_정전기청소포 200매 1박스',
('살림백서 실내건조 세탁세제 2.7L 2개입(드럼용) 대용량 액체세제', '살림백서 실내건조 세탁세제 2.7L 2개입(드럼용) 대용량 액체세제'): '■4WC_SB0101_실내건조 세탁세제 (드럼) 2.7L 2개입',
('살림백서 그린 엑티브 캡슐세제 60개입 액체 캡슐형 세탁세제 드럼용 일반', '살림백서 그린 엑티브 캡슐세제 60개입 액체 캡슐형 세탁세제 드럼용 일반'): '♥3M_SB0098_그린 엑티브 캡슐세제 30개입 2팩',
('살림백서 식기세척기 린스 헹굼보조제 가정용 액상형 1L sk lg 밀레 삼성 호환', '살림백서 식기세척기 린스 헹굼보조제 가정용 액상형 1L sk lg 밀레 삼성 호환'): '♥3M_SB0040_식세기 헹굼보조제 1L 1개입',
('1+1 살림백서 오푼티아&밤부 샴푸 500ml', '향 선택=02.화이트머스크향'): '♡4W_OPS0002_샴푸 (화이트머스크) 500ml 2개입',
('살림백서 베이비&키즈 아기 바디워시 500ml', '옵션 선택=베이비&키즈 아기 바디워시 500ml'): '♡4W_BA0003_아기 바디워시 500ml 1개입',
('살림백서 울세제 중성세제 리필형 1L 3개입', '살림백서 울세제 중성세제 리필형 1L 3개입'): '■4WC_SB0013_울세제 리필형 1L 3개입',
('1+1 살림백서 천연 지수81% 샴푸 500ml 약산성 퍼퓸 비듬 향기좋은 사춘기 청소년 초등학생 지성 두피 정수리냄새 대용량 오푼티아 밤부', '향:유칼립투스향 샴푸 500ml 1+1-1개'): '♡4W_OPS0006_샴푸 (유칼립투스) 500ml 2개입',
('1+1 살림백서 곰팡이제거제 젤 벽 벽지 화장실 베란다 실리콘 창틀 욕실 결로 곰팡이', '옵션선택):1+1+1 살림백서 바르는 곰팡이제거제-1개'): '♥4M_SB0016_바르는 곰팡이제거제 150g 3개입',
('1+1 살림백서 세탁조 클리너 드럼세탁기 청소 통돌이 통세척 8회분 (150g x 8개)', '1+1 살림백서 세탁조 클리너 드럼세탁기 청소 통돌이 통세척 8회분 (150g x 8개)'): '♥4M_SB0003_세탁조 클리너 2박스',
('(단7일! 행사가) 마미바티 아기 젖병세정제 1 + 1  500ml', '(단7일! 행사가) 마미바티 아기 젖병세정제 1 + 1  500ml'): '♥4MC_OMVB1002_마미바티 주방세제 2개입',
('살림백서 주방세제 리필형 1L 3개입 대용량 천연유래98% 리필 과일세척', '향:살림백서 주방세제 리필 1Lx3개(자몽)-1개'): '■4WC_SB0007_주방세제 리필형 1L (자몽향) 3개입',
('1+1 살림백서 천연 지수81% 샴푸 500ml 약산성 퍼퓸 비듬 향기좋은 사춘기 청소년 초등학생 지성 두피 정수리냄새 대용량 오푼티아 밤부', '향:베이비파우더향 샴푸 500ml 1+1-1개'): '♡4W_OPS0001_샴푸 (베이비파우더) 500ml 2개입',
('살림백서 하루세번 구취케어 치약 100g 3개입 입냄새 구취제거 잇몸', '사이즈:01.하루세번 구취케어 치약 100g 3개입-2개'): '♥3M_SB0154_구취케어 치약 3개입',
('살림백서 토탄수 대용량 클렌징폼 500ml', '제품선택=살림백서 토탄수 클렌징폼 500ml'): '♡4W_RH_OPR031_R허브 토탄수 클렌징폼 500ml 1개입',
('마미바티 아기 섬유유연제 1800ml 2개입', '옵션 선택): 1)마미바티 아기섬유유연제 2개입'): '♥3M_OMVA1005_마미바티 섬유유연제 2개입',
('살림백서 핸드워시 라이스&허브 버블 500ml 손세정제 그린포레스트향 4개 거품비누 핸드솝 대용량', '살림백서 핸드워시 라이스&허브 버블 500ml 손세정제 그린포레스트향 4개 거품비누 핸드솝 대용량'): '■4W_RH_OPR025_R허브 핸드워시 그린포레스트 500ml 4개입',
('살림백서 퍼퓸 섬유유연제 리필형 2L X 3개입', '옵션선택)=1.퍼퓸 섬유유연제 리필 2Lx3개 일랑일랑'): '■4WC_SBB002_섬유유연제 2L 리필 (일랑일랑) 3개입',
('살림백서 이염방지시트 3팩 총 60매 세탁티슈 빨래 흰색옷이염 물빠짐방지제 이염제거', '살림백서 이염방지시트 3팩 총 60매 세탁티슈 빨래 흰색옷이염 물빠짐방지제 이염제거'): '♥4WC_SB0065_이염방지시트 3팩',
('살림백서 딥퍼퓸 섬유유연제 1L X 3개입 5종향 실내건조 향기좋은 대용량', '향:5.딥퍼퓸 섬유유연제 코지 블랑 1L 3개-1개'): '■3MC_SB0132_딥퍼퓸 섬유유연제 (코지블랑) 1L 3개입',
('1+1 살림백서 라이스 앤 허브 샴푸 1L 약산성 퍼퓸 향기좋은 정수리 냄새 청소년 사춘기', '향:RH1.딥그린시더우드 샴푸 (1+1)-1개'): '♡4W_RH_OPR002_R허브 샴푸 딥그린시더우드 1L 2개입',
('살림백서 소프트 퍼퓸 건조기시트 80매 건조기섬유유연제 드라이시트 3종향', '향:1)퍼퓸 건조기 섬유유연제 80매(블랑향)-1개'): '♥4MC_SB0072_건조기 섬유유연제 (블랑향) 2박스',
('살림백서 하루세번 구취케어 치약 100g 3개입 입냄새 구취제거 잇몸', '사이즈:03.하루세번 시그니처 치약 150g 3개입-1개 (+2000원)'): '♥3M_SB0046_치약 3개입',
('살림백서 소프트 퍼퓸 건조기시트 40매 건조기섬유유연제 드라이시트 3종향', '향:1)퍼퓸 건조기 섬유유연제 40매(블랑향)-1개'): '♥4MC_SB0071_건조기 섬유유연제 (블랑향) 1박스',
('1+1 살림백서 디퓨저 라이스&허브 200ml 아로마 실내 방향제 화장실 인테리어 대용량', '향:3)1+1 살림백서 디퓨저 블랙체리-11개'): '♡4W_RH_SB0117_R허브 디퓨저 (블랙체리) 200ml 2개입',
('살림백서 주방세제 500ml (청귤향) 4개입', '살림백서 주방세제 500ml (청귤향) 4개입'): '♥3MC_SB0011_주방세제 500ml (청귤향) 4개입',
('1+1 살림백서 라이스 앤 허브 샴푸 1L 약산성 퍼퓸 향기좋은 정수리 냄새 청소년 사춘기', '향:RH2.그린포레스트 샴푸 (1+1)-1개'): '♡4W_RH_OPR004_R허브 샴푸 그린포레스트 1L 2개입',
('마미바티 베이킹소다 2kg', '마미바티 베이킹소다 2kg'): '♥3M_OMVA1010_마미바티 베이킹 용기 1개입',
('마미바티 과탄산소다 표백제 2kg 용기형 과탄산소다나트륨 대용량', '마미바티 과탄산소다 표백제 2kg 용기형 과탄산소다나트륨 대용량'): '♥3M_OMVA1008_마미바티 과탄산 용기 1개입',
('1+1 살림백서 딥퍼퓸 섬유탈취제 500ml 7종 섬유향수 정전기방지 스프레이 드레스', '향:2)1+1 베이비파우더 500ml-1개'): '■♥4WC_SB0140_딥퍼퓸 섬유탈취제 (베이비파우더) 500ml 2개입',
('살림백서 하루세번 구취케어 치약 100g 3개입 입냄새 구취제거 잇몸', '사이즈:03.하루세번 시그니처 치약 150g 3개입-2개 (+4000원)'): '♥3M_SB0046_치약 3개입',
('쉼표365 촉촉 비데티슈 비데물티슈 60매 10팩 화장실 화장실용 비데용 물에녹는', '쉼표365 촉촉 비데티슈 비데물티슈 60매 10팩 화장실 화장실용 비데용 물에녹는'): '■#1M_OCOA2001_쉼표 비데 물티슈 60매 10개입 ♣',
('1+1 살림백서 차량용 방향제 디퓨저 80ml 6종향 고급 자동차 차량 명품', '구성품:1+1 살림백서 차량방향제 클린솝-1개'): '♡4W_RH_SB0150_차량방향제 (클린솝) 80ml 2개입',
('살림백서 2단계 초 미세모 딥클린 칫솔 20개입 초극세모 임산부 대용량', '살림백서 2단계 초 미세모 딥클린 칫솔 20개입 초극세모 임산부 대용량'): '★4M_SBDC001_딥클린 칫솔 20개입 1박스',
('살림백서 밀대걸레 물걸레 청소포 모음전', '옵션선택)=03.살림백서 물걸레청소포 5팩 - 청소포만'): '■4MC_SB_3PL_6/ 살림백서 물걸레청소포 5팩 ',
('살림백서 신발탈취제 신발냄새제거제 2BOX 6팩 운동화 구두 신발장 냄새제거', '살림백서 신발탈취제 신발냄새제거제 2BOX 6팩 운동화 구두 신발장 냄새제거'): '■4WC_SB0052_슈즈케어 2박스',
('1+1 살림백서 오푼티아&밤부 바디워시 500ml', '향 선택=03. 베베머스크향'): '♡4W_OPB0003_바디워시 (베베머스크) 500ml 2개입',
('쉼표365 인덕션클리너 하이라이트 세정제 가스레인지 가스렌지 청소 세척 세제 후드 기름때 전용', '쉼표365 인덕션클리너 하이라이트 세정제 가스레인지 가스렌지 청소 세척 세제 후드 기름때 전용'): '#3M_OCOB1006_인덕션클리너 350ml 1개입 ♣',
('1+1 살림백서 천연 지수81% 샴푸 500ml 약산성 퍼퓸 비듬 향기좋은 사춘기 청소년 초등학생 지성 두피 정수리냄새 대용량 오푼티아 밤부', '향:일랑일랑 샴푸 500ml 1+1-1개'): '♡4W_OPS0004_샴푸 (일랑일랑) 500ml 2개입',
('살림백서 밀대걸레 물걸레 청소포 모음전', '옵션선택)=01.살림백서 막대걸래(본품) - 밀대만'): '■4MC_SB_3PL_8/ 살림백서 막대걸레 1개',
('살림백서 차량용 방향제 디퓨저 80ml 6종향', '제품선택=차량용 디퓨저 80ml 2개입, 향기선택=클린솝'): '♡4W_RH_SB0150_차량방향제 (클린솝) 80ml 2개입',
('1+1 살림백서 차량용 방향제 디퓨저 80ml 6종향 고급 자동차 차량 명품', '구성품:1+1 살림백서 차량방향제 베이비파우더-1개'): '♡4W_RH_SB0152_차량방향제 (베이비파우더) 80ml 2개입',
('살림백서 일회용 수세미 주방 60매 2롤', '색상:일회용 수세미 (브라운)-1개'): '♥3M_SB0042_수세미 브라운 2롤',
('살림백서 실내건조 세탁세제 2.7L 4개입(일반용) 대용량 액체세제', '살림백서 실내건조 세탁세제 2.7L 4개입(일반용) 대용량 액체세제'): '■4WC_SB0100_실내 건조용 세탁세제 (일반) 2.7L 4개입',
('(단하루!) 살림백서1+1 세탁세제 섬유유연제 주방세제 베이킹', '옵션명 1:04_살림백서 베이킹소다3kg 리필형 1+1-1개 (+3000원)'): '■4WC_SBF002_베이킹소다 3kg 2개입',
('(단하루!) 살림백서1+1 세탁세제 섬유유연제 주방세제 베이킹', '옵션명 1:05_살림백서 구연산1kg 리필형 1+1-2개'): '■4WC_SBG002_구연산 1kg 2개입',
('1+1 살림백서 라이스&허브 디퓨저 200ml (아로마 실내 방향제)', '향 선택=3)살림백서 디퓨저 1+1 블랙체리'): '♡4W_RH_SB0117_R허브 디퓨저 (블랙체리) 200ml 2개입',
('살림백서 일회용 수세미 주방 60매 2롤', '색상:일회용 수세미 (화이트)-1개'): '♥3M_SB0043_수세미 화이트 2롤',
('1+1 살림백서 베이킹소다 3kg', '1+1 살림백서 베이킹소다 3kg'): '■4WC_SBF002_베이킹소다 3kg 2개입',
('1+1 살림백서 오푼티아&밤부 바디로션 500ml', '바디로션 향 선택1=03.바디로션 베베머스크, 바디로션 향 선택2=03.바디로션 베베머스크'): '♡4W_OPBL004_바디로션 (베베머스크향) 500ml 2개입',
('1+1 살림백서 오푼티아&밤부 바디로션 500ml', '바디로션 향 선택1=02.바디로션 화이트머스크, 바디로션 향 선택2=02.바디로션 화이트머스크'): '♡4W_OPBL002_바디로션 (화이트머스크향) 500ml 2개입',
('1+1 살림백서 오푼티아&밤부 바디워시 500ml', '향 선택=02.화이트머스크향'): '♡4W_OPB0002_바디워시 (화이트머스크) 500ml 2개입',
('살림백서 쌀겨 대용량 클렌징폼 500ml', '제품선택=살림백서 쌀겨 클렌징폼 500ml'): '♡4W_RH_OPR029_R허브 쌀겨 클렌징폼 500ml 1개입',
('1+1 살림백서 디퓨저 라이스&허브 200ml 아로마 실내 방향제 화장실 인테리어 대용량', '향:4)1+1 살림백서 디퓨저 포레스트가든-2개'): '♡4W_RH_SB0115_R허브 디퓨저 (포레스트가든) 200ml 2개입',
('살림백서 제습백서 젤리 제습제 모음전 일반형 옷장걸이형', '구성품:2) 젤리 제습백서 틈새형 4개입 4팩(16개입)-1개'): '♥3M_SB0069_젤리 제습백서 일반형 (1BOX)',
('1+1 살림백서 디퓨저 라이스&허브 200ml 아로마 실내 방향제 화장실 인테리어 대용량', '향:2)1+1 살림백서 디퓨저 클린솝-1개'): '♡4W_RH_SB0116_R허브 디퓨저 (클린솝) 200ml 2개입',
('살림백서 옷걸이형 제습제 옷장형 습기제거제 150g 10개입', '살림백서 옷걸이형 제습제 옷장형 습기제거제 150g 10개입'): '■3M_SB0084_옷걸이형 제습제 (비드형) 10개입',
('살림백서 딥클린 주방세제 750ml 3개입 대용량 주방세제 설거지', '향:1)딥클린 주방세제 그린허브 750ml 3개-2개'): '★4M_SBDC007_딥클린 주방세제 그린허브 750ml 3개입',
('1+1 살림백서 식기세척기 세제 액상형 1L 식세기 sk lg 밀레 삼성 호환', '1+1 살림백서 식기세척기 세제 액상형 1L 식세기 sk lg 밀레 삼성 호환'): '♥3M_SB0037_식세기 세제 1L (액체) 2개입',
('살림백서 주방세제 리필형 1L 3개입 대용량 천연유래98% 리필 과일세척', '향:살림백서 주방세제 리필 1Lx3개(청귤)-2개'): '■4WC_SB0010_주방세제 리필형 1L (청귤향) 3개입',
('1+1 살림백서 대용량 바디로션 라이스&허브 1000ml 향좋은 고보습 촉촉한 끈적임없는', '향:RH2.그린포레스트 바디로션 2개-1개'): '향:RH2.그린포레스트 바디로션 2개-1개',
('살림백서 딥퍼퓸 섬유유연제 1L X 3개입 5종향 실내건조 향기좋은 대용량', '향:2.딥퍼퓸 섬유유연제 베이비파우더 1L 3개-1개'): '♥3MC_SB0129_딥퍼퓸 섬유유연제 (베이비파우더) 1L 3개입',
('살림백서 소프트 퍼퓸 건조기 섬유유연제 40매 3종향', '향 선택)=2)퍼퓸 건조기 섬유유연제 40매(러브미향)'): '♥4MC_SB0073_건조기 섬유유연제 (러브미향) 1박스',
('살림백서 소프트 퍼퓸 건조기 섬유유연제 40매 3종향', '향 선택)=1)퍼퓸 건조기 섬유유연제 40매(블랑향)'): '♥4MC_SB0071_건조기 섬유유연제 (블랑향) 1박스',
('1+1 마미바티 아기세제 아기섬유유연제', '선택1: 01_1+1 마미바티 아기세제1800ml'): '♥3M_OMVA1002_마미바티 아기세제 2개입',
('살림백서 실내건조 세탁세제 2.7L x 4개입 대용량 액체세제 드럼용 일반용', '상품선택=2)실내건조 세탁세제 2.7L 4개입(드럼용)'): '■4WC_SB0102_실내건조 세탁세제 (드럼) 2.7L 4개입',
('1+1 살림백서 오푼티아&밤부 바디로션 500ml', '바디로션 향 선택1=03.바디로션 베베머스크, 바디로션 향 선택2=03.바디로션 베베머스크'): '♡4W_OPBL004_바디로션 베베머스크향 500ml 2개입',
('1+1 살림백서 오푼티아&밤부 바디로션 500ml', '바디로션 향 선택1=02.바디로션 화이트머스크, 바디로션 향 선택2=04.바디로션 체리블라썸'): '♡4W_OPBL106_바디로션 화이트머스크+체리블라썸 500ml',
('살림백서 먼지털이 스틱+리필10개입 먼지떨이 세트 떨이개 제거기 털이개', '살림백서 먼지털이 스틱+리필10개입 먼지떨이 세트 떨이개 제거기 털이개'): '♥4MC_SB0121_먼지떨이 스틱+리필10개입 SET',
('살림백서 실내건조 세탁세제 2.7L 4개입(드럼용) 대용량 액체세제', '살림백서 실내건조 세탁세제 2.7L 4개입(드럼용) 대용량 액체세제'): '■4WC_SB0102_실내건조 세탁세제 (드럼) 2.7L 4개입',
('살림백서 천연유래98% 액체 세탁세제 2L 6개입 중성세제 드럼용 일반용 겸용', '살림백서 천연유래98% 액체 세탁세제 2L 6개입 중성세제 드럼용 일반용 겸용'): 'SB_A_3/ 세탁세제 2L (리필) 6개입',
('살림백서 뽑아쓰는 키친타올 100매 9팩', '살림백서 뽑아쓰는 키친타올 100매 9팩'): 'SB_3PL_2/ 살림백서 (뽑)키친타월 1BOX',
('살림백서 보들보들 각티슈 미용티슈 250매 6개입', '살림백서 보들보들 각티슈 미용티슈 250매 6개입'): '3PL_3/ 살림백서 미용티슈 1BOX',
('살림백서 베이비&키즈 아기세제 아기세탁세제 클린솝 1950ml 2개입', '살림백서 베이비&키즈 아기세제 아기세탁세제 클린솝 1950ml 2개입'): '♥3M_BA0007_아기유연제 1950ml 2개입',
('살림백서 변기세정제 40g 10개입', '살림백서 변기세정제 40g 10개입'): '♥4MC_SB0050_변기세정제 1박스',
('마미바티 과탄산소다 2kg 산소계표백제', '마미바티 과탄산소다 2kg 산소계표백제'): '♥3M_OMVA1008_마미바티 과탄산 용기 1개입',
('1+1 살림백서 배수구클리너 150g 4개 (총 8회분)', '1+1 살림백서 배수구클리너 150g 4개 (총 8회분)'): '♥4M_SB0001_배수구 클리너 2박스',
('1+1 살림백서 뿌리는 신발탈취제 200ml 신발장 신발냄새제거제', '1+1 살림백서 뿌리는 신발탈취제 200ml 신발장 신발냄새제거제'): '■4WC_SB0052_슈즈케어 2박스',
('살림백서 제습제 습기제거제 520mlx24개', '살림백서 제습제 습기제거제 520mlx24개'): 'SB_3PL_1/ 살림백서 제습제 1BOX',
('살림백서 흰옷 표백제 1kg X 2 EA', '살림백서 흰옷 표백제 1kg X 2 EA'): '살림백서 흰옷 표백제 1kg X 2 EA',
('1+1 살림백서 주방세제 청귤향 1L(펌프포함)', '1+1 살림백서 주방세제 청귤향 1L(펌프포함)'): '♥4WC_SB0009_주방세제 1L (청귤향) 2개입',
('1+1 살림백서 산소계 표백제 1kg 흰옷 얼룩제거', '1+1 살림백서 산소계 표백제 1kg 흰옷 얼룩제거'): '■4WC_SBH001_흰옷표백제 1kg 2개입',
('살림백서 만능 얼룩제거제 리필형 1L 2개입', '살림백서 만능 얼룩제거제 리필형 1L 2개입'): '♥3MC_SB0020_얼룩제거제 500ml 2개입',
('살림백서 제습제 520ml 24개입 1box 옷장 화장실 신발장', '살림백서 제습제 520ml 24개입 1box 옷장 화장실 신발장'): 'SB_3PL_1/ 살림백서 제습제 1BOX',
('1+1 살림백서 탈모 샴푸 엑티브B7 맥주효모 앤 비오틴 1000ml', '1+1 살림백서 탈모 샴푸 엑티브B7 맥주효모 앤 비오틴 1000ml'): '■4WC_OPP004_탈모샴푸 1L 2개입',
('피부백서 순면 엠보 화장솜 1000매 대용량 (500매 2입)', '피부백서 순면 엠보 화장솜 1000매 대용량 (500매 2입)'): '♡4W_PB0013_엠보 소프트 화장솜 2팩 (총 1,000매)',
('살림백서 변기세정제 (40g X 10개)','살림백서 변기세정제 (40g X 10개)'):'♥4MC_SB0050_변기세정제 1박스',
('살림백서 옷걸이형 제습제 옷장 옷걸이 습기제거제 150g 10개입','옵션 선택=옷걸이형  습기제거제 150g 10개입'):'■3M_SB0084_옷걸이형 제습제 (비드형) 10개입',
('살림백서 변기세정제 40g x 10개','살림백서 변기세정제 40g x 10개'):'♥4MC_SB0050_변기세정제 1박스',
('살림백서 대용량 주방세제 라이스&허브 4L+750ml 석류향 설거지 리필 세트 과일세척', '향:2)주방세제 4L+750ml 스위트허브향-1개'): '■4WC_SB0111_R허브 주방세제 (스위트허브) 4L 1개입+750ml 1개입',
('살림백서 딥클린 코코넛 주방 수세미 6개 1세트 설거지 셀룰로오스 제로웨이스트', '살림백서 딥클린 코코넛 주방 수세미 6개 1세트 설거지 셀룰로오스 제로웨이스트'): '★4M_SBDC013_딥클린 코코넛 주방 수세미 6개입 1박스',
('살림백서 베이비&키즈 아기 엉덩이 클렌저 300ml x 2개', '옵션 선택=베이비&키즈 아기 엉덩이 클렌저 300ml x 2개'): '♡4W_BA0009_아기 엉덩이 세정제 300ml 2개입',
('1+1 살림백서 오푼티아&밤부 트리트먼트 500ml', '향 선택=02.화이트머스크향'): '♡4W_OPT0002_트리 (화이트머스크) 500ml 2개입',
('1+1 살림백서 오푼티아&밤부 트리트먼트 500ml', '향 선택=03.체리블라썸향'): '♡4W_OPT0005_트리 (체리블라썸) 500ml 2개입',
('살림백서 라이스 앤 허브 샴푸 1L 약산성 퍼퓸 향기좋은 정수리 냄새 청소년 사춘기', '향:RH1.딥그린시더우드 샴푸 1개-1개'): '♡4W_RH_OPR001_R허브 샴푸 딥그린시더우드 1L 1개입',
('1+1 살림백서 라이스 앤 허브 바디워시 1L 대용량 퍼퓸 향좋은 샤워젤 바디클렌져', '향:RH2.그린포레스트 바디워시 (1+1)-1개'): '♡4W_RH_OPR010_R허브 바디워시 딥그린시더우드 1L 2개입',
('[살림백서] 헤어 바디 베스트 모음전 샴푸/바디워시/바디로션/핸드크림/핸드워시 외 오푼티아 라이스앤허브 탈모 아크네', '제품 선택:1+1 살림백서 대용량 바디로션 라이스앤허브 1000ml,향 선택:RH1.딥그린시더우드-1개 (+2600원)'): '♡4W_RH_OPR014_R허브 바디로션 딥그린시더우드 1L 2개입',
('살림백서 대용량 디퓨저 500ml X 2개 선물세트 아로마 실내 방향제 포레스트가든 화장실', '향:1+1 살림백서 디퓨저 대용량 500ml 비자림숲-1개'): '♡4W_RH_SB0187_R허브 디퓨저 (비자림숲) 500ml 2개입 ',
('살림백서 하루세번 천연 유래 치약 150g 3개입', '옵션 선택=하루세번 천연 유래 치약 150g 3개입'): '♥3M_SB0046_치약 3개입',
('살림백서 하루세번 잇몸케어 치약 100g 3개입', '살림백서 하루세번 잇몸케어 치약 100g 3개입'): '♥3M_SB0155_잇몸케어 치약 3개입',
('1+1 살림백서 얼룩제거제', '1+1 살림백서 얼룩제거제'): '♥4M_SB0020_얼룩제거제 500ml 2개입',
('1+1 살림백서 섬유탈취제 500ml 섬유향수 드레스퍼퓸 스프레이 탈취99% 정전기방지', '향:1+1 살림백서 섬유탈취제 500ml(블랑)-1개'): '♥4M_SB0080_섬유탈취제 (코지블랑) 500ml 2개입',
('살림백서 1+1 딥퍼퓸 섬유탈취제 500ml 7종', '향 선택=1+1 화이트머스크 500ml'): '■♥4WC_SB0139_딥퍼퓸 섬유탈취제 (화이트머스크) 500ml 2개입',
('살림백서 2단계 초 미세모 딥클린 칫솔 20개입', '살림백서 2단계 초 미세모 딥클린 칫솔 20개입'): '★4M_SBDC001_딥클린 칫솔 20개입 1박스',
('[살림백서] 헤어 바디 베스트 모음전 샴푸/바디워시/바디로션/핸드크림/핸드워시 외 오푼티아 라이스앤허브 탈모 아크네', '제품 선택:1+1 살림백서 오푼티아 샴푸 500ml,향 선택:04.일랑일랑 샴푸 1+1 500ml-1개'): '♡4W_OPS0004_샴푸 (일랑일랑) 500ml 2개입',
('캡슐 세탁조 클리너 통돌이 드럼 겸용 6개입', '캡슐 세탁조 클리너 6개입 X 1세트'): '♥3M_SB0170_캡슐 세탁조 클리너 6개입 1개',














  
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
