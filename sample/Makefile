# Markdown to Word変換用Makefile

# 設定
INPUT_MD = sample.md
OUTPUT_DOCX = sample.docx
TITLE = SampleApp
SUBTITLE =
DOCTYPE = ユーザーマニュアル
VERSION = 1.0.0
DATE = 2025年11月28日
DEPT = システム開発部
DOCNUM = SAMPLE-APP-100
LOGO = logo.png
COMPANY = サンプル株式会社

# デフォルトターゲット
.PHONY: all
all: $(OUTPUT_DOCX)

# Word文書の生成
$(OUTPUT_DOCX): $(INPUT_MD) ../md2mdocx.js
	bun ../md2mdocx.js $(INPUT_MD) $(OUTPUT_DOCX) \
		--title "$(TITLE)" \
		--subtitle "$(SUBTITLE)" \
		--doctype "$(DOCTYPE)" \
		--version "$(VERSION)" \
		--date "$(DATE)" \
		--dept "$(DEPT)" \
		--docnum "$(DOCNUM)" \
		--logo "$(LOGO)" \
		--company "$(COMPANY)"

# クリーン
.PHONY: clean
clean:
	rm -f $(OUTPUT_DOCX)

# 再生成
.PHONY: rebuild
rebuild: clean all

# ヘルプ
.PHONY: help
help:
	@echo "使用方法:"
	@echo "  make          - USER_MANUAL.docxを生成"
	@echo "  make clean    - 生成したdocxを削除"
	@echo "  make rebuild  - クリーンしてから再生成"
	@echo ""
	@echo "カスタマイズ例:"
	@echo "  make TITLE=\"製品名\" VERSION=\"2.0.0\""
	@echo "  make SUBTITLE=\"取扱説明書\" DOCTYPE=\"ユーザーガイド\""
	@echo "  make LOGO=\"company_logo.png\""
	@echo "  make COMPANY=\"株式会社ABC\""
