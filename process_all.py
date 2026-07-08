"""
批量处理脚本 —— 一键识别+生成

用法（3种，任选其一）：

  1. 把招标文件拖到 input/ 文件夹，直接运行：
     python process_all.py

  2. 指定任意文件夹：
     python process_all.py ./我的招标文件/

  3. 使用AI增强（自动生成高质量技术方案）：
     python process_all.py --api-key YOUR_OPENAI_API_KEY

  4. 只识别不生成（预览模式）：
     python process_all.py --scan-only

  5. 配合公司配置文件：
     python process_all.py --company ./我的公司信息.json
"""
import os
import sys
import argparse
import json
from datetime import datetime

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.bid_recognizer import BidFileRecognizer
from src.bid_document_parser import BiddingDocumentParser
from src.bid_response_generator import BidResponseGenerator
from src.document_output import generate_docx


def main():
    parser = argparse.ArgumentParser(
        description='投标文件批量生成器 —— 放入招标文件，一键生成投标文件',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    
    parser.add_argument('folder', nargs='?', default=None,
                        help='招标文件所在文件夹（默认: ./input/）')
    parser.add_argument('--scan-only', '-s', action='store_true',
                        help='仅扫描识别，不生成（预览模式）')
    parser.add_argument('--api-key', '-k', type=str, default=None,
                        help='OpenAI API Key，启用AI智能生成')
    parser.add_argument('--base-url', type=str, default=None,
                        help='OpenAI代理地址')
    parser.add_argument('--company', '-c', type=str, default=None,
                        help='公司信息配置文件路径')
    parser.add_argument('--bid-info', type=str, default=None,
                        help='投标信息配置文件（含报价等）')
    parser.add_argument('--output', '-o', type=str, default=None,
                        help='输出目录（默认: ./output/）')

    args = parser.parse_args()

    # ---- 初始化 ----
    project_root = os.path.dirname(os.path.abspath(__file__))
    input_dir = args.folder or os.path.join(project_root, 'input')
    output_dir = args.output or os.path.join(project_root, 'output')
    
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    # ---- 步骤1: 扫描识别 ----
    print("\n" + "=" * 60)
    print("  📋 步骤 1/4：扫描目录，识别招标文件...")
    print("=" * 60)
    
    recognizer = BidFileRecognizer(input_dir)
    result = recognizer.print_report()

    bidding_files = result['bidding_docs']
    
    if not bidding_files:
        print("\n⚠️  未发现招标文件！")
        print(f"   请将招标文件放入: {input_dir}")
        print(f"   支持的格式: PDF, DOCX, TXT, XLSX 等")
        print(f"\n   也可以指定文件夹: python process_all.py ./你的文件夹")
        
        if result['unrecognized']:
            print(f"\n   💡 提示: 有 {len(result['unrecognized'])} 个无法识别的文件，")
            print(f"      你可以直接把它们移到 input/ 重试，或用 --scan-only 预览内容")
        
        return

    if args.scan_only:
        print("\n✅ 扫描完成（预览模式，未生成文件）")
        return

    # ---- 步骤2: 加载配置 ----
    print("=" * 60)
    print("  📋 步骤 2/4：加载配置...")
    print("=" * 60)

    # 公司信息
    company_profile = args.company
    if company_profile and os.path.exists(company_profile):
        print(f"  公司信息: {company_profile}")
    else:
        company_profile = None
        print(f"  公司信息: 使用默认 (data/company_profile.json)")

    # 投标信息（报价等）
    bid_extra = {}
    if args.bid_info and os.path.exists(args.bid_info):
        with open(args.bid_info, 'r', encoding='utf-8') as f:
            config = json.load(f)
            bid_extra = config.get('bid_info', {})
        print(f"  投标信息: {args.bid_info}")

    bid_extra['submission_date'] = bid_extra.get('submission_date', 
                                                   datetime.now().strftime('%Y年%m月%d日'))

    # ---- 步骤3: 逐个生成 ----
    print("\n" + "=" * 60)
    print(f"  📋 步骤 3/4：生成投标文件（共 {len(bidding_files)} 个）...")
    print("=" * 60)

    generator = BidResponseGenerator(company_profile_path=company_profile)
    
    if args.api_key:
        generator.enable_llm(args.api_key, args.base_url)
        print("  ✨ AI增强模式已启用")

    results = []
    
    for i, file_info in enumerate(bidding_files, 1):
        filepath = file_info['path']
        filename = file_info['filename']
        
        print(f"\n  [{i}/{len(bidding_files)}] 处理: {filename}")
        print(f"  {'-' * 50}")
        
        try:
            # 解析招标文件
            print(f"    ⏳ 解析中...")
            parser_bid = BiddingDocumentParser()
            bidding_info = parser_bid.parse(filepath)

            project = bidding_info.get('project_info', {})
            proj_name = project.get('project_name', '未知项目')
            proj_id   = project.get('project_id', '未识别')
            bid_date  = project.get('bid_opening_date', '未识别')
            
            print(f"    项目名称: {proj_name}")
            print(f"    项目编号: {proj_id}")
            print(f"    开标日期: {bid_date}")
            print(f"    招标人:   {project.get('tenderee_name', '未识别')}")
            print(f"    文档章节: {len(bidding_info.get('required_documents', []))} 个")
            print(f"    技术要求: {len(bidding_info.get('requirements', {}).get('technical', []))} 条")

            # 生成投标内容
            print(f"    ⏳ 生成投标内容...")
            bid_doc = generator.generate(bidding_info, bid_extra)
            print(f"    已生成 {len(bid_doc)} 个章节: {', '.join(ch.get('title', k)[:15] for k, ch in list(bid_doc.items())[:6])}")

            # 输出DOCX
            print(f"    ⏳ 生成DOCX...")
            output_path = generate_docx(bid_doc, {**project, **bid_extra})
            
            # 移动到统一输出目录
            import shutil
            final_path = os.path.join(output_dir, os.path.basename(output_path))
            shutil.move(output_path, final_path)
            
            print(f"    ✅ 完成 → {final_path}")
            results.append({'status': 'success', 'file': filename, 'output': final_path})
            
        except Exception as e:
            print(f"    ❌ 失败: {e}")
            results.append({'status': 'failed', 'file': filename, 'error': str(e)})

    # ---- 步骤4: 汇总报告 ----
    print("\n\n" + "=" * 60)
    print("  📋 步骤 4/4：生成汇总报告")
    print("=" * 60)

    success_count = sum(1 for r in results if r['status'] == 'success')
    fail_count = sum(1 for r in results if r['status'] == 'failed')

    print(f"\n  🎉 批量处理完成！")
    print(f"     成功: {success_count} 个")
    print(f"     失败: {fail_count} 个")
    print(f"     输出: {output_dir}")

    if success_count > 0:
        print(f"\n  📁 生成的文件：")
        for r in results:
            if r['status'] == 'success':
                print(f"     ✓ {r['output']}")

    if fail_count > 0:
        print(f"\n  ⚠️ 处理失败的文件：")
        for r in results:
            if r['status'] == 'failed':
                print(f"     ✗ {r['file']} - {r['error']}")

    # 生成JSON报告
    report_path = os.path.join(output_dir, f'生成报告_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json')
    with open(report_path, 'w', encoding='utf-8') as f:
        json.dump({
            'generated_at': datetime.now().isoformat(),
            'input_dir': input_dir,
            'output_dir': output_dir,
            'total': len(bidding_files),
            'success': success_count,
            'failed': fail_count,
            'details': results,
        }, f, ensure_ascii=False, indent=2)
    print(f"\n  📄 详细报告: {report_path}")
    
    print("\n" + "=" * 60 + "\n")

    # 尝试打开输出文件夹
    try:
        os.startfile(output_dir)
    except:
        pass


if __name__ == "__main__":
    main()
