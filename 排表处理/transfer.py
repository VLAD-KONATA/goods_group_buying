import os
import pandas as pd
import argparse

def trading_results(cn,good):
    try:
        cn=cn.split('（')[0].strip() if '（' in cn else cn.split('(')[0].strip()
        result=cn.split('（')[1].strip() if '（' in cn else cn.split('(')[1].strip()
        result=cn.split('）')[0].strip() if '）' in cn else cn.split(')')[0].strip()
    except Exception as e:
        print(f"处理数据时出错：{e}")
        return None
    good=good+result
    return cn,good

def process_table(args):
    TABLE_TYPE=["排表","导入表"]
    IGNORE_LIST=["窗 (哀)","小瓶子+秀（猫）哀","Ling（黑羽盗一）","风止 (兰)","远忧（新兰花栗鼠）",
                "好心路过草莓鹿（停谷版）","傲娇 (小哀)","绯初 (推景零)","无尽夏 (哀妈)","念念 (兰，和叶)",
                "懿 (哀)","澍(佳佳)","小音 (推琴酒雪莉)","www (新兰)","秋稀 (柯妈)","熊抱 (新兰)"]
    DEBUG=args.debug
    series=args.series
    storage=args.storage
    # 解析路径
    base_dir = os.path.join(args.dir,storage,TABLE_TYPE[0],series+".xlsx")  # 假设table_path在排表/萌/下
    import_dir = os.path.join(args.dir,storage,TABLE_TYPE[1],series)
    os.makedirs(import_dir, exist_ok=True)
    NO_UPDATE_LIST=os.listdir(import_dir)

    # 读取xlsx
    xls = pd.ExcelFile(base_dir)
    for sheet_name in xls.sheet_names:
        # 如果导入表中已经存在同名文件，则跳过处理该sheet)(调试模式时除外)
        if sheet_name+".xlsx" in NO_UPDATE_LIST and (not DEBUG):
            continue
        print(f"Processing sheet: {sheet_name}")
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        if df.empty:
            continue

        table_name = sheet_name  # 表名即sheet名

        # 判断类型
        if df.shape[0] < 2:
            continue

        if df.iloc[1, 0] == '种类':
            # 调价类型（拼盒）
            if df.shape[0] < 3:
                continue
            goods = df.iloc[1, 1:].dropna().tolist()
            data_rows = df.iloc[3:, :]
            result = []
            for idx, row in data_rows.iterrows():
                seq = row.iloc[0]
                if pd.isna(seq):
                    continue
                for i, cn in enumerate(row.iloc[1:], 1):
                    if pd.isna(cn) or i-1 >= len(goods):
                        continue
                    good = goods[i-1]
                    #盲抽检测
                    if ('（' in cn or '(' in cn) and cn not in IGNORE_LIST:
                        cn, good = trading_results(cn, good)
                    result.append({'cn': str(cn).strip(), '谷名': good})
            if result:
                print(f"录入{len(result)}条数据")
                result_df = pd.DataFrame(result)
                grouped = result_df.groupby(['cn', '谷名']).size().reset_index(name='数量')
                grouped['系列'] = series
                grouped['表名'] = table_name
                grouped['肾/国际状态'] = ''
                grouped['囤货'] = storage
                grouped['排发状态'] = '不可排发'
                # 重新排序列
                grouped = grouped[['cn', '系列', '表名', '谷名', '数量', '肾/国际状态', '囤货', '排发状态']]
                output_path = os.path.join(import_dir, f'{sheet_name}.xlsx')
                grouped.to_excel(output_path, index=False, header=True)
                print(f"{series}/{sheet_name} completed {len(grouped)}.\n")
        elif df.iloc[1, 0].lower() == 'cn':
            # 单领类型
            data_rows = df.iloc[2:, :]
            result = []
            current_cn = None
            for idx, row in data_rows.iterrows():
                cn = row.iloc[0]
                good_raw = row.iloc[1] if df.shape[1] > 1 else None
                if pd.notna(cn):
                    current_cn = str(cn).strip()
                if pd.isna(good_raw):
                    continue
                good = str(good_raw).strip()
                quantity = 1
                result.append({
                    'cn': current_cn,
                    '系列': series,
                    '表名': table_name,
                    '谷名': good,
                    '数量': quantity,
                    '肾/国际状态': '',
                    '囤货': storage,
                    '排发状态': '不可排发'
                })
            if result:
                print(f"录入{len(result)}条数据")
                result_df = pd.DataFrame(result)
                result_df = result_df[['cn', '系列', '表名', '谷名', '数量', '肾/国际状态', '囤货', '排发状态']]
                output_path = os.path.join(import_dir, f'{sheet_name}.xlsx')
                result_df.to_excel(output_path, index=False, header=True)
                print(f"{series}/{sheet_name} completed {len(result_df)}.\n")
        # 其他类型跳过

if __name__ == '__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument("--series",type=str,required=True,help='系列')
    parser.add_argument("--dir",type=str,default=".\\example_file_structure\\",help='目录')
    parser.add_argument("--storage",type=str,default='萌',help='囤货')
    parser.add_argument("--debug",action='store_true',help='调试模式')
    args=parser.parse_args()
    process_table(args)