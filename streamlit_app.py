import streamlit as st
import pandas as pd
import io
import uuid

st.set_page_config(
    page_title="Excel数据提取工具",
    page_icon="📊",
    layout="centered"
)

def extract_data_from_excel(uploaded_file):
    """从Excel文件中提取SKU、ASIN、FNSKU数据"""
    try:
        # 读取Excel文件
        df = pd.read_excel(uploaded_file)
        
        # 如果只有一列，可能需要转换
        if len(df.columns) == 1:
            data = df.iloc[:, 0].tolist()
        else:
            # 如果有多列，查找包含数据的列
            data = []
            for col in df.columns:
                data.extend(df[col].dropna().astype(str).tolist())
        
        # 提取SKU、ASIN、FNSKU
        extracted_data = []
        
        for i, item in enumerate(data):
            if str(item).strip() == 'ASIN' and i+1 < len(data):
                asin_value = str(data[i+1]).strip()
                
                # 查找对应的SKU
                sku_value = ""
                for j in range(i+2, min(i+8, len(data))):
                    if str(data[j]).strip() == 'SKU' and j+1 < len(data):
                        sku_value = str(data[j+1]).strip()
                        break
                
                # 查找对应的FNSKU
                fnsku_value = ""
                for j in range(i+2, min(i+15, len(data))):
                    if str(data[j]).strip() == 'FNSKU' and j+1 < len(data):
                        fnsku_value = str(data[j+1]).strip()
                        break
                
                # 添加到结果中
                extracted_data.append({
                    'ASIN': asin_value,
                    'SKU': sku_value,
                    'FNSKU': fnsku_value
                })
        
        return extracted_data
        
    except Exception as e:
        st.error(f"提取数据时出错: {str(e)}")
        return []

def main():
    st.title("📊 Excel数据提取工具")
    st.markdown("**上传Excel文件，自动提取SKU、ASIN和FNSKU信息**")
    
    # 文件上传
    uploaded_file = st.file_uploader(
        "选择Excel文件",
        type=['xlsx', 'xls'],
        help="支持 .xlsx 和 .xls 格式"
    )
    
    if uploaded_file is not None:
        # 显示文件信息
        st.info(f"📁 已选择文件: {uploaded_file.name} ({uploaded_file.size / 1024 / 1024:.2f} MB)")
        
        # 处理按钮
        if st.button("🚀 提取数据", type="primary"):
            with st.spinner("正在处理文件，请稍候..."):
                # 提取数据
                extracted_data = extract_data_from_excel(uploaded_file)
                
                if extracted_data:
                    st.success(f"✅ 提取成功！共找到 **{len(extracted_data)}** 个商品信息")
                    
                    # 显示数据预览
                    df_result = pd.DataFrame(extracted_data)
                    st.subheader("📋 数据预览")
                    st.dataframe(df_result, use_container_width=True)
                    
                    # 下载按钮
                    csv_buffer = io.StringIO()
                    df_result.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
                    csv_data = csv_buffer.getvalue()
                    
                    st.download_button(
                        label="📥 下载CSV文件",
                        data=csv_data,
                        file_name=f"extracted_data_{uuid.uuid4().hex[:8]}.csv",
                        mime="text/csv",
                        type="secondary"
                    )
                    
                    # 统计信息
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("总商品数", len(extracted_data))
                    with col2:
                        valid_sku = sum(1 for item in extracted_data if item['SKU'])
                        st.metric("有效SKU", valid_sku)
                    with col3:
                        valid_fnsku = sum(1 for item in extracted_data if item['FNSKU'])
                        st.metric("有效FNSKU", valid_fnsku)
                
                else:
                    st.error("❌ 没有找到有效的SKU、ASIN或FNSKU数据")
                    st.info("请确保Excel文件包含正确格式的数据")
    
    # 使用说明
    with st.expander("📖 使用说明"):
        st.markdown("""
        ### 📋 支持的数据格式
        - Excel文件中应包含 **ASIN**、**SKU**、**FNSKU** 关键词
        - 数据可以在任意列中，工具会自动识别
        
        ### 🔄 提取逻辑
        1. 找到"ASIN"关键词，获取下一行的值
        2. 在ASIN附近查找"SKU"关键词和对应值
        3. 在ASIN附近查找"FNSKU"关键词和对应值
        
        ### 📊 输出格式
        - CSV格式文件
        - 包含ASIN、SKU、FNSKU三列数据
        """)
    
    # 页脚
    st.markdown("---")
    st.markdown("*🛠️ 数据提取工具 | 快速处理Excel文件*")

if __name__ == "__main__":
    main()