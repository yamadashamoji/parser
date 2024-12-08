import sys
import glob
import os
import codecs
import csv
from xml.etree.ElementTree import parse
from pathlib import Path
import PyPDF2

def get_pdf_page_count(directory):
    """
    指定されたディレクトリ内の単一のPDFファイルのページ数を取得する関数

    Args:
        directory (str): PDFファイルが格納されているディレクトリのパス

    Returns:
        int: PDFのページ数 (PDFが見つからない場合は0)
    """
    # ディレクトリ内のファイルリストを取得
    files = [f for f in os.listdir(directory) if f.lower().endswith('.pdf')]

    # PDFファイルが1つでない場合は0を返す
    if len(files) != 1:
        print(f"エラー: ディレクトリ {directory} には1つのPDFファイルが必要です。現在のファイル数: {len(files)}")
        return 0

    # PDFファイルのフルパスを作成
    pdf_path = os.path.join(directory, files[0])

    try:
        # PDFファイルを開く
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            # ページ数を取得
            page_count = len(pdf_reader.pages)
            return page_count

    except Exception as e:
        print(f"PDFファイルの読み取り中にエラーが発生しました: {e}")
        return 0

def xml_to_csv(ipt, opt):
    p = Path(ipt)
    xml_path = list(p.glob('**/*.xml'))
    os.makedirs(opt, exist_ok=True)
    os.chdir(opt)

    # 名前空間の定義
    namespaces = {
        'xsd': 'http://www.w3.org/2001/XMLSchema',
        'jppat': 'http://www.jpo.go.jp/standards/XMLSchema/ST96/JPPatent',
        'com': 'http://www.wipo.int/standards/XMLSchema/ST96/Common',
        'pat': 'http://www.wipo.int/standards/XMLSchema/ST96/Patent',
        'jpcom': 'http://www.jpo.go.jp/standards/XMLSchema/ST96/JPCommon',
        'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    }

    for xml_file in xml_path:
        try:
            print(f"Processing file: {xml_file}")
            
            tree = parse(xml_file)
            elem = tree.getroot()

            # データ抽出関数
            def safe_get_text(xpath, join_text=""):
                try:
                    element = elem.find(xpath, namespaces=namespaces)
                    if element is not None:
                        if join_text:
                            return join_text.join(element.itertext()).replace('\n', '').strip()
                        return element.text.strip() if element.text else ""
                    return ""
                except Exception as e:
                    print(f"Error extracting {xpath}: {str(e)}")
                    return ""

            # 公報種別の取得
            publication_status = safe_get_text('.//pat:PlainLanguageDesignationText')
            
            # 公開特許公報(A)または公表特許公報(A)の場合のみ処理を続行
            if publication_status in ["公開特許公報(A)", "公表特許公報(A)"]:
                # CSV名の決定
                csv_name = elem.findtext('.//com:ApplicationNumber/com:ApplicationNumberText', namespaces=namespaces)
                if not csv_name:
                    csv_name = xml_file.stem
                
                csv_path = Path(opt) / f"{csv_name}.csv"
                
                # PDFページ数の取得（XMLと同じディレクトリを想定）
                pdf_directory = xml_file.parent
                page_count = get_pdf_page_count(pdf_directory)
                
                with open(csv_path, 'a', encoding='utf-8', newline='') as csv_file:
                    writer = csv.writer(csv_file)

                    # データ抽出（すべて相対パスに修正）
                    data = [
                        "Z",  # 既存のデータベースと帳尻合わせをするため
                        "Y",  # 既存のデータベースと帳尻合わせをするため
                        "種別なし",  # 既存のデータベースと帳尻合わせをするため
                        safe_get_text('.//pat:PublicationNumber'),
                        safe_get_text('.//com:PublicationDate').replace("-", ""),
                        safe_get_text('.//com:ApplicationNumber/com:ApplicationNumberText'),
                        safe_get_text('.//jppat:ApplicationIdentification/pat:FilingDate').replace("-", ""),
                        safe_get_text('.//pat:InventionTitle'),
                        safe_get_text('.//jppat:IPCClassification/pat:MainClassification'),
                        safe_get_text('.//jppat:ClaimTotalQuantitySet/pat:ClaimTotalQuantity'),
                        
                        # PDFページ数を追加
                        page_count,
                        
                        # 以下、既存のコードを続ける
                        #FI ここはappendでスペース入れないとスペース入れてくれない
                        safe_get_text('.//jppat:UnexaminedPatentPublicationBibliographicData/jppat:NationalClassification', "      "),
                        #テーマコード
                        safe_get_text('.//jppat:ThemeCodeInformationBag/jppat:ThemeCodeInformation', ""),
                        # Fターム（一部Fタームの記載の無い公開特許公報（A) があるのでエラーを吐き出す)
                        safe_get_text('.//jppat:FtermInformationBag/jppat:FtermInformation', ""),
                        
                        # 出願人情報
                        safe_get_text('.//jppat:Applicant[1]/com:PartyIdentifier'),                        
                        safe_get_text('.//jppat:Applicant[1]/jpcom:Contact/com:Name/com:EntityName'),  
                        safe_get_text('.//jppat:Applicant[1]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:Applicant[2]/com:PartyIdentifier'),                        
                        safe_get_text('.//jppat:Applicant[2]/jpcom:Contact/com:Name/com:EntityName'),  
                        safe_get_text('.//jppat:Applicant[2]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:Applicant[3]/com:PartyIdentifier'),                        
                        safe_get_text('.//jppat:Applicant[3]/jpcom:Contact/com:Name/com:EntityName'),  
                        safe_get_text('.//jppat:Applicant[3]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:Applicant[4]/com:PartyIdentifier'),                        
                        safe_get_text('.//jppat:Applicant[4]/jpcom:Contact/com:Name/com:EntityName'),  
                        safe_get_text('.//jppat:Applicant[4]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:Applicant[5]/com:PartyIdentifier'),                        
                        safe_get_text('.//jppat:Applicant[5]/jpcom:Contact/com:Name/com:EntityName'),  
                        safe_get_text('.//jppat:Applicant[5]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:Applicant[6]/com:PartyIdentifier'),                        
                        safe_get_text('.//jppat:Applicant[6]/jpcom:Contact/com:Name/com:EntityName'),  
                        safe_get_text('.//jppat:Applicant[6]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:Applicant[7]/com:PartyIdentifier'),                        
                        safe_get_text('.//jppat:Applicant[7]/jpcom:Contact/com:Name/com:EntityName'),  
                        safe_get_text('.//jppat:Applicant[7]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:Applicant[8]/com:PartyIdentifier'),                        
                        safe_get_text('.//jppat:Applicant[8]/jpcom:Contact/com:Name/com:EntityName'),  
                        safe_get_text('.//jppat:Applicant[8]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:Applicant[9]/com:PartyIdentifier'),                        
                        safe_get_text('.//jppat:Applicant[9]/jpcom:Contact/com:Name/com:EntityName'),  
                        safe_get_text('.//jppat:Applicant[9]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:Applicant[10]/com:PartyIdentifier'),                        
                        safe_get_text('.//jppat:Applicant[10]/jpcom:Contact/com:Name/com:EntityName'),  
                        safe_get_text('.//jppat:Applicant[10]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:Applicant[11]/com:PartyIdentifier'),                        
                        safe_get_text('.//jppat:Applicant[11]/jpcom:Contact/com:Name/com:EntityName'),  
                        safe_get_text('.//jppat:Applicant[11]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:Applicant[12]/com:PartyIdentifier'),                        
                        safe_get_text('.//jppat:Applicant[12]/jpcom:Contact/com:Name/com:EntityName'),  
                        safe_get_text('.//jppat:Applicant[12]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),

                        # 代理人情報
                        safe_get_text('.//jppat:RegisteredPractitioner[1]/pat:RegisteredPractitionerRegistrationNumber'),
                        safe_get_text('.//jppat:RegisteredPractitioner[1]/jpcom:Contact/com:Name/com:EntityName'),
                        safe_get_text('.//jppat:RegisteredPractitioner[2]/pat:RegisteredPractitionerRegistrationNumber'),
                        safe_get_text('.//jppat:RegisteredPractitioner[2]/jpcom:Contact/com:Name/com:EntityName'),
                        safe_get_text('.//jppat:RegisteredPractitioner[3]/pat:RegisteredPractitionerRegistrationNumber'),
                        safe_get_text('.//jppat:RegisteredPractitioner[3]/jpcom:Contact/com:Name/com:EntityName'),
                        safe_get_text('.//jppat:RegisteredPractitioner[4]/pat:RegisteredPractitionerRegistrationNumber'),
                        safe_get_text('.//jppat:RegisteredPractitioner[4]/jpcom:Contact/com:Name/com:EntityName'),
                        safe_get_text('.//jppat:RegisteredPractitioner[5]/pat:RegisteredPractitionerRegistrationNumber'),
                        safe_get_text('.//jppat:RegisteredPractitioner[5]/jpcom:Contact/com:Name/com:EntityName'),
                        safe_get_text('.//jppat:RegisteredPractitioner[6]/pat:RegisteredPractitionerRegistrationNumber'),
                        safe_get_text('.//jppat:RegisteredPractitioner[6]/jpcom:Contact/com:Name/com:EntityName'),
                        safe_get_text('.//jppat:RegisteredPractitioner[7]/pat:RegisteredPractitionerRegistrationNumber'),
                        safe_get_text('.//jppat:RegisteredPractitioner[7]/jpcom:Contact/com:Name/com:EntityName'),
                        safe_get_text('.//jppat:RegisteredPractitioner[8]/pat:RegisteredPractitionerRegistrationNumber'),
                        safe_get_text('.//jppat:RegisteredPractitioner[8]/jpcom:Contact/com:Name/com:EntityName'),
                        safe_get_text('.//jppat:RegisteredPractitioner[9]/pat:RegisteredPractitionerRegistrationNumber'),
                        safe_get_text('.//jppat:RegisteredPractitioner[9]/jpcom:Contact/com:Name/com:EntityName'),
                        safe_get_text('.//jppat:RegisteredPractitioner[10]/pat:RegisteredPractitionerRegistrationNumber'),
                        safe_get_text('.//jppat:RegisteredPractitioner[10]/jpcom:Contact/com:Name/com:EntityName'),
                        safe_get_text('.//jppat:RegisteredPractitioner[11]/pat:RegisteredPractitionerRegistrationNumber'),
                        safe_get_text('.//jppat:RegisteredPractitioner[11]/jpcom:Contact/com:Name/com:EntityName'),
                        safe_get_text('.//jppat:RegisteredPractitioner[12]/pat:RegisteredPractitionerRegistrationNumber'),
                        safe_get_text('.//jppat:RegisteredPractitioner[12]/jpcom:Contact/com:Name/com:EntityName'),
                        
                        # 発明者情報
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[1]/jpcom:Contact/com:Name/com:EntityName'),                        
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[1]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[2]/jpcom:Contact/com:Name/com:EntityName'),                        
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[2]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[3]/jpcom:Contact/com:Name/com:EntityName'),                        
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[3]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[4]/jpcom:Contact/com:Name/com:EntityName'),                        
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[4]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[5]/jpcom:Contact/com:Name/com:EntityName'),                        
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[5]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[6]/jpcom:Contact/com:Name/com:EntityName'),                        
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[6]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[7]/jpcom:Contact/com:Name/com:EntityName'),                        
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[7]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[8]/jpcom:Contact/com:Name/com:EntityName'),                        
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[8]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[9]/jpcom:Contact/com:Name/com:EntityName'),                        
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[9]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[10]/jpcom:Contact/com:Name/com:EntityName'),                        
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[10]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[11]/jpcom:Contact/com:Name/com:EntityName'),                        
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[11]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[12]/jpcom:Contact/com:Name/com:EntityName'),                        
                        safe_get_text('.//jppat:InventorBag/jppat:Inventor[12]/jpcom:Contact/com:PostalAddressBag/com:PostalAddress/com:PostalAddressText'),

                        #要約【課題】＋【解決手段】＋【選択図】
                        safe_get_text('.//pat:Abstract/com:P', "    "),
                        # 請求項（すべて）
                        safe_get_text('.//pat:Claims/pat:Claim/pat:ClaimText', "    "),
                        # 技術分野（すべて）
                        safe_get_text('.//jppat:Description/pat:TechnicalField/com:P', "    "),
                        # 背景技術（すべて）
                        safe_get_text('.//jppat:Description/pat:BackgroundArt/com:P', "    "),
                        # 特許文献（すべて）
                        safe_get_text('.//com:CitationBag/com:PatentCitationBag/com:P/com:PatentCitation/com:PatentCitationText', "    "),
                        # 非特許文献（すべて）
                        safe_get_text('.//com:CitationBag/com:NPLCitationBag/com:P/com:NPLCitation/com:NPLCitationText', "    "),
                        # 発明が解決しようとする課題
                        safe_get_text('.//pat:InventionSummary/pat:TechnicalProblem/com:P', "    "),
                        # 発明を解決するための手段
                        safe_get_text('.//pat:InventionSummary/pat:TechnicalSolution/com:P', "    "),
                        # 発明の効果
                        safe_get_text('.//pat:InventionSummary/pat:AdvantageousEffects/com:P', "    "),
                        # 発明を実施するための形態
                        safe_get_text('.//pat:EmbodimentDescription/com:P', "    "),
                        # 産業利用上の可能性
                        safe_get_text('.//pat:IndustrialApplicability/com:P', "    "),
                        # 図面の簡単な説明
                        safe_get_text('.//pat:DrawingDescription/com:P/com:FigureReference', "    ")
                    ]

                    writer.writerow(data)
                    print(f"Successfully processed {xml_file}")
            else:
                print(f"Skipping {xml_file} - Publication status: {publication_status}")

        except Exception as e:
            print(f"Error processing {xml_file}: {str(e)}")
            continue

if __name__ == "__main__":
    ipt = sys.argv[1] if len(sys.argv) > 1 else R"" # Zip解凍後のファイル
    opt = sys.argv[2] if len(sys.argv) > 2 else R"" # csv変換後のファイル
    
    xml_to_csv(ipt, opt)