"""
批量处理脚本 v2 —— 严格按招标文件第七章格式生成响应文件

用法:
  1. 把招标文件(.docx)放到 input/ 文件夹，直接运行：
     python process_all.py

  2. 仅预览识别结果：
     python process_all.py --scan-only

  3. 指定公司信息：
     python process_all.py --company ./我的公司.json
"""
import os
import sys
import argparse
import json
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.bid_recognizer import BidFileRecognizer
from src.bid_document_parser import BiddingDocumentParser
from src.response_template_extractor import ResponseTemplateExtractor


def main():
    parser = argparse.ArgumentParser(description='投标响应文件批量生成器 v2')
    parser.add_argument('folder', nargs='?', default=None, help='招标文件所在文件夹（默认: ./input/）')
    parser.add_argument('--scan-only', '-s', action='store_true', help='仅扫描识别')
    parser.add_argument('--company', '-c', type=str, default=None, help='公司信息配置文件')
    parser.add_argument('--output', '-o', type=str, default=None, help='输出目录')

    args = parser.parse_args()

    project_root = os.path.dirname(os.path.abspath(__file__))
    input_dir = args.folder or os.path.join(project_root, 'input')
    output_dir = args.output or os.path.join(project_root, 'output')
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    # ==================== 步骤1: 扫描识别 ====================
    print("\n" + "=" * 60)
    print("  步骤 1/3：扫描目录，识别招标文件...")
    print("=" * 60)

    recognizer = BidFileRecognizer(input_dir)
    scan_result = recognizer.print_report()
    bidding_files = scan_result['bidding_docs']

    if not bidding_files:
        print("\n  未发现招标文件！请将 .docx 招标文件放入 input/ 文件夹")
        if scan_result['unrecognized']:
            print(f"  提示: 有 {len(scan_result['unrecognized'])} 个未识别文件")
        return

    if args.scan_only:
        print("\n  扫描完成（预览模式，未生成文件）")
        return

    # ==================== 步骤2: 加载公司信息 ====================
    print("\n" + "=" * 60)
    print("  步骤 2/3：加载公司信息...")
    print("=" * 60)

    company_data = _load_company(args.company, project_root)
    print(f"  公司名称: {company_data.get('bidder_name', '未设置')}")
    print(f"  法定代表人: {company_data.get('legal_representative', '未设置')}")

    # ==================== 步骤3: 逐个生成 ====================
    print("\n" + "=" * 60)
    print(f"  步骤 3/3：生成响应文件（共 {len(bidding_files)} 个）...")
    print("=" * 60)

    extractor = ResponseTemplateExtractor()
    results = []

    for i, file_info in enumerate(bidding_files, 1):
        filepath = file_info['path']
        filename = file_info['filename']

        print(f"\n  [{i}/{len(bidding_files)}] {filename}")
        print(f"  {'-' * 50}")

        try:
            # 只处理 DOCX 文件（需要原始格式）
            if not filename.lower().endswith('.docx'):
                print(f"    跳过: 非 DOCX 文件，无法提取原始格式")
                results.append({'status': 'skipped', 'file': filename, 'reason': '非DOCX'})
                continue

            # ---- 从第一章提取项目信息 ----
            print("    [解析] 解析第一章 磋商邀请...")
            doc_parser = BiddingDocumentParser()
            bidding_info = doc_parser.parse(filepath)
            project = bidding_info.get('project_info', {})

            proj_name = project.get('project_name', project.get('project_name_original', '未知项目'))
            proj_id = project.get('project_id', '未识别')
            bid_date = project.get('bid_opening_date', '未识别')
            agency = project.get('agency_name', '')
            tenderer = project.get('tenderee_name', '')

            print(f"    项目名称: {proj_name}")
            print(f"    项目编号: {proj_id}")
            print(f"    开标日期: {bid_date}")
            if tenderer:
                print(f"    采购人:   {tenderer}")
            if agency:
                print(f"    代理机构: {agency}")

            # ---- 准备填充数据 ----
            fill_data = {
                'project_name': proj_name,
                'project_id': proj_id,
                'bid_opening_date': bid_date,
                'bidder_name': company_data.get('bidder_name', ''),
                'legal_representative': company_data.get('legal_representative', ''),
                'authorized_person': company_data.get('authorized_person', ''),
                'agency_name': agency or tenderer or '贵单位',
                'address': company_data.get('bidder_address', ''),
                'phone': company_data.get('bidder_phone', ''),
                'zip_code': company_data.get('bidder_zip_code', ''),
                'fax': company_data.get('bidder_fax', ''),
                'package_no': '1',
            }

            # ---- 生成响应文件 ----
            print("    [生成] 按招标文件第七章格式生成...")
            output_path = extractor.generate(filepath, fill_data)

            # 移动到统一输出目录
            import shutil
            final_path = os.path.join(output_dir, os.path.basename(output_path))
            if os.path.abspath(output_path) != os.path.abspath(final_path):
                shutil.move(output_path, final_path)

            print(f"    [成功] 完成 -> {os.path.basename(final_path)}")
            results.append({
                'status': 'success',
                'file': filename,
                'project_name': proj_name,
                'project_id': proj_id,
                'output': final_path
            })

        except Exception as e:
            import traceback
            print(f"    [失败] 失败: {e}")
            traceback.print_exc()
            results.append({'status': 'failed', 'file': filename, 'error': str(e)})

    # ==================== 汇总报告 ====================
    print("\n\n" + "=" * 60)
    print("  汇总报告")
    print("=" * 60)

    success = sum(1 for r in results if r['status'] == 'success')
    failed = sum(1 for r in results if r['status'] == 'failed')
    skipped = sum(1 for r in results if r['status'] == 'skipped')

    print(f"\n  处理完成！成功: {success}  失败: {failed}  跳过: {skipped}")
    print(f"  输出目录: {output_dir}")

    if success > 0:
        print(f"\n  生成的文件：")
        for r in results:
            if r['status'] == 'success':
                print(f"    [OK] {os.path.basename(r['output'])}")
                print(f"      项目: {r['project_name']} ({r['project_id']})")

    if failed > 0:
        print(f"\n  失败的文件：")
        for r in results:
            if r['status'] == 'failed':
                print(f"    [FAIL] {r['file']} - {r['error']}")

    # 保存JSON报告
    report_path = os.path.join(output_dir, f'生成报告_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json')
    with open(report_path, 'w', encoding='utf-8') as f:
        json.dump({
            'time': datetime.now().isoformat(),
            'total': len(bidding_files),
            'success': success,
            'failed': failed,
            'skipped': skipped,
            'details': results,
        }, f, ensure_ascii=False, indent=2)
    print(f"\n  详细报告: {report_path}")

    # 尝试打开输出文件夹
    try:
        os.startfile(output_dir)
    except:
        pass

    print("\n" + "=" * 60 + "\n")


def _load_company(config_path: str, project_root: str) -> dict:
    """加载公司信息"""
    if config_path and os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data.get('company', data)

    default_path = os.path.join(project_root, 'data', 'company_profile.json')
    if os.path.exists(default_path):
        with open(default_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data.get('company', data)

    print("  警告：未找到公司配置文件，将使用占位符")
    return {}


if __name__ == "__main__":
    main()
