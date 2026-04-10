import streamlit as st
import pandas as pd
import numpy as np
from scipy import stats
from io import BytesIO
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings
warnings.filterwarnings('ignore')

# 设置页面
st.set_page_config(
    page_title="Hard Carbon CV Analysis",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🔬 Sodium-ion Battery Hard Carbon Anode CV Analysis")
st.markdown("### Capacitive Contribution and Sodium Storage Mechanism Analysis")
st.markdown("---")

# 侧边栏参数设置
with st.sidebar:
    st.header("⚙️ Parameters")
    
    # 电极参数
    st.subheader("📐 Electrode Parameters")
    A = st.number_input("Electrode Area A (cm²)", value=0.5026, format="%.4f",
                        help="Geometric or electrochemically active area of the electrode")
    
    st.markdown("---")
    
    # 电压区间设置
    st.subheader("📊 Voltage Region Division")
    st.markdown("Two characteristic regions based on hard carbon sodium storage properties:")
    
    col1, col2 = st.columns(2)
    with col1:
        region1_start = st.number_input("Region I Start (V)", value=0.8, step=0.1, format="%.2f",
                                        help="Adsorption/defect region start", min_value=0.0, max_value=2.5)
        region1_end = st.number_input("Region I End (V)", value=0.1, step=0.05, format="%.2f",
                                      help="Adsorption/defect region end", min_value=0.0, max_value=2.5)
    
    with col2:
        region2_start = st.number_input("Region II Start (V)", value=0.1, step=0.01, format="%.2f",
                                        help="Insertion/pore filling region start", min_value=0.0, max_value=2.5)
        region2_end = st.number_input("Region II End (V)", value=0.01, step=0.01, format="%.2f",
                                      help="Insertion/pore filling region end", min_value=0.0, max_value=2.5)
    
    st.markdown("---")
    
    # 分析选项
    st.subheader("🔧 Analysis Options")
    
    scan_direction = st.radio("Scan Direction", ["Forward Scan (0V → 2.5V)", "Reverse Scan (2.5V → 0V)"], index=1)
    
    st.markdown("---")
    st.caption("📁 Data format: 10 columns, ordered by scan rates 0.1, 0.25, 0.5, 1.0, 2.0 mV/s")
    st.caption("⚠️ Note: Each scan rate contains voltage and current columns")

# 主界面
st.header("📁 Upload Data File")

uploaded_file = st.file_uploader(
    "Select Excel file containing five scan rates data", 
    type=['xlsx', 'xls'],
    help="File should contain 10 columns: 0.1_Ewe_V, 0.1_I_mA, 0.25_Ewe_V, 0.25_I_mA, 0.5_Ewe_V, 0.5_I_mA, 1.0_Ewe_V, 1.0_I_mA, 2.0_Ewe_V, 2.0_I_mA"
)

def correct_b_value(b):
    """Correct b-value to range 0.5-1"""
    if b < 0.5:
        return 0.5
    elif b > 1.0:
        return 1.0
    else:
        return b

def separate_scan_directions(voltage, current):
    """Separate forward and reverse scan data"""
    # Find the minimum voltage point (turning point)
    min_v_idx = np.argmin(voltage)
    
    # Reverse scan: from high to low potential (2.5V → 0V)
    reverse_scan_v = voltage[:min_v_idx+1]
    reverse_scan_i = current[:min_v_idx+1]
    
    # Forward scan: from low to high potential (0V → 2.5V)
    forward_scan_v = voltage[min_v_idx:]
    forward_scan_i = current[min_v_idx:]
    
    return {
        'forward': {'voltage': forward_scan_v, 'current': forward_scan_i},
        'reverse': {'voltage': reverse_scan_v, 'current': reverse_scan_i}
    }

def filter_voltage_range(voltage, current, v_min=0, v_max=0.8):
    """Filter data within specified voltage range"""
    mask = (voltage >= v_min) & (voltage <= v_max)
    return voltage[mask], current[mask]

def create_square_figure(size=1000):
    """Create a square figure with consistent styling - 1:1 aspect ratio"""
    fig = go.Figure()
    fig.update_layout(
        width=size,
        height=size,
        plot_bgcolor='white',
        paper_bgcolor='white',
        font={'size': 48, 'color': 'black', 'family': 'Arial Black'},
        margin=dict(l=120, r=120, t=150, b=120)
    )
    return fig

def style_square_axis(fig, x_title="", y_title="", x_range=None, y_range=None, log_x=False, log_y=False):
    """Apply consistent axis styling for square figures with complete borders"""
    axis_style = {
        'title': {'text': f'<b>{x_title}</b>', 'font': {'size': 48, 'color': 'black', 'family': 'Arial Black'}},
        'tickfont': {'size': 42, 'color': 'black', 'family': 'Arial Black'},
        'showline': True,
        'linewidth': 4,
        'linecolor': 'black',
        'mirror': True,
        'ticks': 'outside',
        'tickwidth': 3,
        'tickcolor': 'black',
        'ticklen': 10,
        'gridcolor': 'lightgray',
        'griddash': 'dot'
    }
    if x_range:
        axis_style['range'] = x_range
    if log_x:
        axis_style['type'] = 'log'
    
    fig.update_xaxes(**axis_style)
    
    yaxis_style = {
        'title': {'text': f'<b>{y_title}</b>', 'font': {'size': 48, 'color': 'black', 'family': 'Arial Black'}},
        'tickfont': {'size': 42, 'color': 'black', 'family': 'Arial Black'},
        'showline': True,
        'linewidth': 4,
        'linecolor': 'black',
        'mirror': True,
        'ticks': 'outside',
        'tickwidth': 3,
        'tickcolor': 'black',
        'ticklen': 10,
        'gridcolor': 'lightgray',
        'griddash': 'dot'
    }
    if y_range:
        yaxis_style['range'] = y_range
    if log_y:
        yaxis_style['type'] = 'log'
    
    fig.update_yaxes(**yaxis_style)
    
    return fig

if uploaded_file is not None:
    try:
        # 读取数据
        df = pd.read_excel(uploaded_file)
        
        # 定义扫速和列名
        scan_rates = [0.1, 0.25, 0.5, 1.0, 2.0]
        expected_columns = [
            '0.1_Ewe_V', '0.1_I_mA',
            '0.25_Ewe_V', '0.25_I_mA',
            '0.5_Ewe_V', '0.5_I_mA',
            '1.0_Ewe_V', '1.0_I_mA',
            '2.0_Ewe_V', '2.0_I_mA'
        ]
        
        # 检查列名
        missing_cols = [col for col in expected_columns if col not in df.columns]
        if missing_cols:
            st.error(f"❌ Missing columns: {missing_cols}")
            st.info("Please ensure column names are correct, e.g., 0.1_Ewe_V, 0.1_I_mA, ...")
        else:
            st.success("✅ File loaded successfully!")
            
            # 显示数据预览
            with st.expander("📋 Data Preview", expanded=False):
                st.dataframe(df.head(10), use_container_width=True)
                st.write(f"Total data points: {len(df)}")
            
            # 检查所有电压列是否一致
            v_common = df['0.1_Ewe_V'].dropna().values
            all_voltage_same = True
            
            for rate in scan_rates[1:]:
                v_check = df[f'{rate}_Ewe_V'].dropna().values
                if len(v_check) != len(v_common) or not np.allclose(v_check, v_common, rtol=1e-5):
                    all_voltage_same = False
                    
            if all_voltage_same:
                st.success("✅ All voltage columns are consistent")
            else:
                st.warning("⚠️ Voltage columns are not completely consistent, but program will continue with respective voltage data")
            
            # 处理每个扫速的数据，分离正扫和反扫
            cv_data = {direction: {} for direction in ['forward', 'reverse']}
            
            for rate in scan_rates:
                v_col = f'{rate}_Ewe_V'
                i_col = f'{rate}_I_mA'
                
                voltage = df[v_col].dropna().values
                current = df[i_col].dropna().values
                
                # 确保数据长度一致
                min_len = min(len(voltage), len(current))
                voltage = voltage[:min_len]
                current = current[:min_len]
                
                # 分离扫描方向
                separated = separate_scan_directions(voltage, current)
                
                for direction in ['forward', 'reverse']:
                    # 筛选0-0.8V范围内的数据
                    v_filtered, i_filtered = filter_voltage_range(
                        separated[direction]['voltage'],
                        separated[direction]['current'],
                        v_min=0,
                        v_max=0.8
                    )
                    
                    cv_data[direction][rate] = {
                        'voltage': v_filtered,
                        'current': i_filtered,
                        'v_Vs': rate / 1000
                    }
            
            # 确定要分析的扫描方向
            current_direction = 'forward' if 'Forward' in scan_direction else 'reverse'
            direction_label = "Forward" if current_direction == 'forward' else "Reverse"
            
            st.info(f"📊 Current analysis: {direction_label} Scan (Display range: 0-0.8V)")
            
            # 创建选项卡
            tab1, tab2, tab3, tab4 = st.tabs([
                "📈 b-value Analysis", 
                "📊 Capacitive Contribution Analysis",
                "🔍 Sodium Storage Mechanism at 0.1V",
                "📋 Comprehensive Report"
            ])
            
            # ==================== Tab 1: b值分析 ====================
            with tab1:
                st.subheader(f"📈 b-value Analysis - {direction_label} Scan")
                
                # 获取所有扫速的电压数据
                v_common = cv_data[current_direction][scan_rates[0]]['voltage']
                
                # 对每个电位点计算b值
                b_values = []
                b_errors = []
                b_r2 = []
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for idx, v in enumerate(v_common):
                    status_text.text(f"Calculating b-value: {idx+1}/{len(v_common)}")
                    
                    # 收集该电位下所有扫速的电流
                    currents = []
                    valid_rates = []
                    
                    for rate in scan_rates:
                        v_data = cv_data[current_direction][rate]['voltage']
                        i_data = cv_data[current_direction][rate]['current']
                        
                        # 找到最接近的电压点
                        closest_idx = np.argmin(np.abs(v_data - v))
                        if np.abs(v_data[closest_idx] - v) < 0.001:
                            currents.append(abs(i_data[closest_idx]))
                            valid_rates.append(rate)
                    
                    # 排除电流太小的点
                    if len(currents) >= 3 and min(currents) > 1e-6:
                        log_v = np.log([cv_data[current_direction][r]['v_Vs'] for r in valid_rates])
                        log_i = np.log(currents)
                        
                        slope, intercept, r_value, p_value, std_err = stats.linregress(log_v, log_i)
                        
                        # 修正b值
                        corrected_b = correct_b_value(slope)
                        
                        b_values.append(corrected_b)
                        b_errors.append(std_err)
                        b_r2.append(r_value**2)
                    else:
                        b_values.append(np.nan)
                        b_errors.append(np.nan)
                        b_r2.append(np.nan)
                    
                    progress_bar.progress((idx + 1) / len(v_common))
                
                status_text.text("b-value calculation completed!")
                progress_bar.empty()
                
                b_values = np.array(b_values)
                b_errors = np.array(b_errors)
                b_r2 = np.array(b_r2)
                
                # 定义两个区域的掩码
                region1_mask = (v_common >= region1_end) & (v_common <= region1_start)
                region2_mask = (v_common >= region2_end) & (v_common <= region2_start)
                
                # 在区域II中找到b值最接近0.5的点
                region2_b = b_values[region2_mask]
                region2_v = v_common[region2_mask]
                region2_r2 = b_r2[region2_mask]
                
                # 找到最接近0.5的b值点（只考虑R² > 0.8的点）
                valid_idx = ~np.isnan(region2_b) & (region2_r2 > 0.8)
                if np.any(valid_idx):
                    valid_b = region2_b[valid_idx]
                    valid_v = region2_v[valid_idx]
                    closest_to_0_5_idx = np.argmin(np.abs(valid_b - 0.5))
                    min_b_voltage = valid_v[closest_to_0_5_idx]
                    min_b_value = valid_b[closest_to_0_5_idx]
                else:
                    min_b_voltage = None
                    min_b_value = None
                
                # ===== 图1: b值图 (正方形1:1) =====
                fig_b = create_square_figure(1000)
                
                # b值曲线
                fig_b.add_trace(
                    go.Scatter(
                        x=v_common, 
                        y=b_values,
                        mode='lines',
                        name='<b>b-value</b>',
                        line=dict(color='red', width=6),
                        error_y=dict(type='data', array=b_errors, visible=True, thickness=3, width=6)
                    )
                )
                
                # 添加R²曲线 - 改为橙色虚线
                fig_b.add_trace(
                    go.Scatter(
                        x=v_common, 
                        y=b_r2,
                        mode='lines',
                        name='<b>R²</b>',
                        line=dict(color='#FF8C00', width=5, dash='dash'),
                        yaxis="y2"
                    )
                )
                
                # 添加参考线 - 标注放在刻度线下方
                fig_b.add_hline(y=0.5, line_dash="dash", line_color="blue", 
                               annotation_text="<b>b=0.5 (diffusion-controlled)</b>", 
                               annotation_position="bottom left",
                               annotation_font=dict(size=42, color='blue', family='Arial Black'))
                fig_b.add_hline(y=1.0, line_dash="dash", line_color="green", 
                               annotation_text="<b>b=1.0 (capacitive-controlled)</b>",
                               annotation_position="bottom left",
                               annotation_font=dict(size=42, color='green', family='Arial Black'))
                
                # 标记两个区域
                fig_b.add_vrect(x0=region1_end, x1=region1_start, 
                               fillcolor="green", opacity=0.15, 
                               annotation_text="<b>Region I<br>Adsorption/Defect</b>",
                               annotation_position="top right",
                               annotation_font=dict(size=42, color='green', family='Arial Black'))
                fig_b.add_vrect(x0=region2_end, x1=region2_start, 
                               fillcolor="red", opacity=0.15,
                               annotation_text="<b>Region II<br>Insertion/Pore filling</b>",
                               annotation_position="top left",
                               annotation_font=dict(size=42, color='red', family='Arial Black'))
                
                # 如果找到了有效点，添加竖直线标注
                if min_b_voltage is not None:
                    fig_b.add_vline(x=min_b_voltage, line_dash="dash", line_color="purple", 
                                   line_width=4,
                                   annotation_text=f"<b>V = {min_b_voltage:.3f}V<br>b = {min_b_value:.3f}</b>",
                                   annotation_position="bottom right",
                                   annotation_font=dict(size=42, color='purple', family='Arial Black'))
                
                # 更新布局 - 添加标题
                fig_b.update_layout(
                    title={
                        'text': '<b>b-value as a Function of Potential</b>',
                        'x': 0.5,
                        'xanchor': 'center',
                        'font': {'size': 54, 'color': 'black', 'family': 'Arial Black'}
                    },
                    xaxis={
                        'title': {'text': '<b>Potential (V)</b>', 'font': {'size': 48, 'color': 'black', 'family': 'Arial Black'}},
                        'tickfont': {'size': 42, 'color': 'black', 'family': 'Arial Black'},
                        'range': [0, 0.8],
                        'showline': True,
                        'linewidth': 4,
                        'linecolor': 'black',
                        'mirror': True,
                        'ticks': 'outside',
                        'tickwidth': 3,
                        'tickcolor': 'black',
                        'ticklen': 10,
                        'gridcolor': 'lightgray',
                        'griddash': 'dot'
                    },
                    yaxis={
                        'title': {'text': '<b>b-value</b>', 'font': {'size': 48, 'color': 'black', 'family': 'Arial Black'}},
                        'tickfont': {'size': 42, 'color': 'black', 'family': 'Arial Black'},
                        'range': [0, 1.2],
                        'showline': True,
                        'linewidth': 4,
                        'linecolor': 'black',
                        'mirror': True,
                        'ticks': 'outside',
                        'tickwidth': 3,
                        'tickcolor': 'black',
                        'ticklen': 10,
                        'gridcolor': 'lightgray',
                        'griddash': 'dot'
                    },
                    yaxis2={
                        'title': {'text': '<b>R²</b>', 'font': {'size': 48, 'color': 'black', 'family': 'Arial Black'}},
                        'tickfont': {'size': 42, 'color': 'black', 'family': 'Arial Black'},
                        'overlaying': 'y',
                        'side': 'right',
                        'range': [0, 1],
                        'showline': True,
                        'linewidth': 4,
                        'linecolor': 'black',
                        'mirror': True,
                        'ticks': 'outside',
                        'tickwidth': 3,
                        'tickcolor': 'black',
                        'ticklen': 10
                    },
                    showlegend=True,
                    legend={
                        'x': 0.85,
                        'y': 0.85,
                        'font': {'size': 42, 'color': 'black', 'family': 'Arial Black'},
                        'bgcolor': 'rgba(255,255,255,0.95)',
                        'bordercolor': 'black',
                        'borderwidth': 3
                    },
                    hovermode='x unified'
                )
                
                st.plotly_chart(fig_b, use_container_width=True)
                
                # 特征区域b值统计
                st.subheader("📊 b-value Statistics in Characteristic Regions")
                
                region_stats = []
                
                # 区域I (只使用R² >= 0.99的数据)
                if np.any(region1_mask):
                    valid_r2_mask = (b_r2 >= 0.99) & region1_mask
                    b_region1 = b_values[valid_r2_mask]
                    b_region1 = b_region1[~np.isnan(b_region1)]
                    r2_region1 = b_r2[valid_r2_mask]
                    r2_region1 = r2_region1[~np.isnan(r2_region1)]
                    
                    if len(b_region1) > 0:
                        mean_b1 = np.mean(b_region1)
                        region_stats.append({
                            'Region': 'Region I (Adsorption/Defect)',
                            'Potential Range': f'{region1_start}-{region1_end} V',
                            'Data Points (R²≥0.99)': len(b_region1),
                            'Mean b-value': f'{mean_b1:.3f}',
                            'Std Dev': f'{np.std(b_region1):.3f}',
                            'Mean R²': f'{np.mean(r2_region1):.3f}',
                            'Dominant Mechanism': 'Capacitive' if mean_b1 > 0.8 else 
                                                  'Diffusion' if mean_b1 < 0.6 else 'Mixed'
                        })
                
                # 区域II (只使用R² >= 0.99的数据)
                if np.any(region2_mask):
                    valid_r2_mask = (b_r2 >= 0.99) & region2_mask
                    b_region2 = b_values[valid_r2_mask]
                    b_region2 = b_region2[~np.isnan(b_region2)]
                    r2_region2 = b_r2[valid_r2_mask]
                    r2_region2 = r2_region2[~np.isnan(r2_region2)]
                    
                    if len(b_region2) > 0:
                        mean_b2 = np.mean(b_region2)
                        region_stats.append({
                            'Region': 'Region II (Insertion/Pore filling)',
                            'Potential Range': f'{region2_start}-{region2_end} V',
                            'Data Points (R²≥0.99)': len(b_region2),
                            'Mean b-value': f'{mean_b2:.3f}',
                            'Std Dev': f'{np.std(b_region2):.3f}',
                            'Mean R²': f'{np.mean(r2_region2):.3f}',
                            'Dominant Mechanism': 'Capacitive' if mean_b2 > 0.8 else 
                                                  'Diffusion' if mean_b2 < 0.6 else 'Mixed'
                        })
                
                st.dataframe(pd.DataFrame(region_stats), use_container_width=True)
            
            # ==================== Tab 2: 电容贡献分析 ====================
            with tab2:
                st.subheader(f"📊 Capacitive Contribution Analysis - {direction_label} Scan")
                
                # 对每个电位点计算k1和k2
                k1_values = []
                k2_values = []
                k_r2 = []
                
                v_Vs = np.array([cv_data[current_direction][rate]['v_Vs'] for rate in scan_rates])
                sqrt_v = np.sqrt(v_Vs)
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for idx, v in enumerate(v_common):
                    status_text.text(f"Calculating contribution: {idx+1}/{len(v_common)}")
                    
                    # 收集该电位下所有扫速的电流
                    currents = []
                    valid_rates = []
                    
                    for rate in scan_rates:
                        v_data = cv_data[current_direction][rate]['voltage']
                        i_data = cv_data[current_direction][rate]['current']
                        
                        closest_idx = np.argmin(np.abs(v_data - v))
                        if np.abs(v_data[closest_idx] - v) < 0.001:
                            currents.append(abs(i_data[closest_idx]) / 1000)
                            valid_rates.append(rate)
                    
                    if len(currents) >= 3 and np.all(np.array(currents) > 1e-8):
                        Y = np.array(currents) / sqrt_v[:len(currents)]
                        X = sqrt_v[:len(currents)]
                        
                        try:
                            slope, intercept, r_value, _, _ = stats.linregress(X, Y)
                            k1_values.append(slope)
                            k2_values.append(intercept)
                            k_r2.append(r_value**2)
                        except:
                            k1_values.append(np.nan)
                            k2_values.append(np.nan)
                            k_r2.append(np.nan)
                    else:
                        k1_values.append(np.nan)
                        k2_values.append(np.nan)
                        k_r2.append(np.nan)
                    
                    progress_bar.progress((idx + 1) / len(v_common))
                
                status_text.text("Contribution calculation completed!")
                progress_bar.empty()
                
                k1_values = np.array(k1_values)
                k2_values = np.array(k2_values)
                k_r2 = np.array(k_r2)
                
                # 定义两个区域的掩码
                region1_mask = (v_common >= region1_end) & (v_common <= region1_start)
                region2_mask = (v_common >= region2_end) & (v_common <= region2_start)
                
                # 计算每个扫速下的贡献率
                scan_contributions = []
                total_cap_ratios = []
                total_diff_ratios = []
                
                for rate in scan_rates:
                    scan_v = rate / 1000
                    sqrt_scan = np.sqrt(scan_v)
                    
                    i_total = k1_values * scan_v + k2_values * sqrt_scan
                    i_cap = k1_values * scan_v
                    
                    cap_ratio = np.where(i_total > 0, (i_cap / i_total) * 100, np.nan)
                    
                    # 计算两个区域的平均贡献率
                    if np.any(region1_mask):
                        cap1 = np.nanmean(cap_ratio[region1_mask])
                    else:
                        cap1 = np.nan
                    
                    if np.any(region2_mask):
                        cap2 = np.nanmean(cap_ratio[region2_mask])
                    else:
                        cap2 = np.nan
                    
                    # 计算总的电容贡献率
                    total_cap = np.nanmean(cap_ratio)
                    total_diff = 100 - total_cap
                    
                    total_cap_ratios.append(total_cap)
                    total_diff_ratios.append(total_diff)
                    
                    scan_contributions.append({
                        'Scan Rate (mV/s)': rate,
                        'Region I Capacitive (%)': f'{cap1:.1f}' if not np.isnan(cap1) else 'N/A',
                        'Region I Diffusive (%)': f'{100-cap1:.1f}' if not np.isnan(cap1) else 'N/A',
                        'Region II Capacitive (%)': f'{cap2:.1f}' if not np.isnan(cap2) else 'N/A',
                        'Region II Diffusive (%)': f'{100-cap2:.1f}' if not np.isnan(cap2) else 'N/A',
                        'Total Capacitive (%)': f'{total_cap:.1f}' if not np.isnan(total_cap) else 'N/A',
                        'Total Diffusive (%)': f'{total_diff:.1f}' if not np.isnan(total_diff) else 'N/A'
                    })
                
                # ===== 图2: 总贡献率柱状图 (正方形1:1) =====
                fig_total_contrib = create_square_figure(1000)
                
                rates_display = [f'{r} mV/s' for r in scan_rates]
                
                fig_total_contrib.add_trace(go.Bar(
                    name='<b>Capacitive</b>',
                    x=rates_display,
                    y=total_cap_ratios,
                    marker_color='#2E86C1',
                    text=[f'<b>{c:.1f}%</b>' for c in total_cap_ratios],
                    textposition='inside',
                    textfont={'size': 38, 'color': 'white', 'family': 'Arial Black'},
                    width=0.4,
                    insidetextanchor='middle'
                ))
                
                fig_total_contrib.add_trace(go.Bar(
                    name='<b>Diffusive</b>',
                    x=rates_display,
                    y=total_diff_ratios,
                    marker_color='#E67E22',
                    text=[f'<b>{d:.1f}%</b>' for d in total_diff_ratios],
                    textposition='inside',
                    textfont={'size': 38, 'color': 'white', 'family': 'Arial Black'},
                    width=0.4,
                    insidetextanchor='middle'
                ))
                
                fig_total_contrib.update_layout(
                    title={
                        'text': '<b>Total Contributions at Different Scan Rates (0-0.8V)</b>',
                        'x': 0.5,
                        'xanchor': 'center',
                        'font': {'size': 54, 'color': 'black', 'family': 'Arial Black'}
                    },
                    xaxis={
                        'title': {'text': '<b>Scan Rate (mV/s)</b>', 'font': {'size': 48, 'color': 'black', 'family': 'Arial Black'}},
                        'tickfont': {'size': 42, 'color': 'black', 'family': 'Arial Black'},
                        'showline': True, 
                        'linewidth': 4, 
                        'linecolor': 'black', 
                        'mirror': True,
                        'ticks': 'outside',
                        'tickwidth': 3,
                        'tickcolor': 'black',
                        'ticklen': 10
                    },
                    yaxis={
                        'title': {'text': '<b>Contribution (%)</b>', 'font': {'size': 48, 'color': 'black', 'family': 'Arial Black'}},
                        'tickfont': {'size': 42, 'color': 'black', 'family': 'Arial Black'},
                        'range': [0, 100], 
                        'showline': True, 
                        'linewidth': 4, 
                        'linecolor': 'black', 
                        'mirror': True,
                        'ticks': 'outside',
                        'tickwidth': 3,
                        'tickcolor': 'black',
                        'ticklen': 10
                    },
                    barmode='stack',
                    showlegend=False
                )
                
                st.plotly_chart(fig_total_contrib, use_container_width=True)
                
                # 显示所有扫速的贡献率
                st.subheader("📊 Contribution Summary at Different Scan Rates")
                contrib_df = pd.DataFrame(scan_contributions)
                st.dataframe(contrib_df, use_container_width=True)
            
            # ==================== Tab 3: 分区储钠机制（三个独立的图）====================
            with tab3:
                st.subheader(f"🔍 Sodium Storage Mechanism Analysis at 0.1V - {direction_label} Scan")
                
                # 分析0.1V分界点
                boundary_v = 0.1
                
                # 找到最近的电位点
                v_idx = np.argmin(np.abs(v_common - boundary_v))
                actual_v = v_common[v_idx]
                
                st.info(f"Analysis potential point: {actual_v:.3f} V (closest to 0.1V)")
                
                # 收集该电位下各扫速的电流
                currents = []
                for rate in scan_rates:
                    v_data = cv_data[current_direction][rate]['voltage']
                    i_data = cv_data[current_direction][rate]['current']
                    closest_idx = np.argmin(np.abs(v_data - actual_v))
                    currents.append(abs(i_data[closest_idx]))
                
                currents_abs = np.array(currents)
                
                # ===== 图3: i vs v 关系 (正方形1:1) =====
                fig1 = create_square_figure(1000)
                
                log_v = np.log([cv_data[current_direction][rate]['v_Vs'] for rate in scan_rates])
                log_i = np.log(currents_abs)
                
                slope, intercept, r_value, _, _ = stats.linregress(log_v, log_i)
                b = correct_b_value(slope)
                
                # 散点图
                fig1.add_trace(
                    go.Scatter(
                        x=[cv_data[current_direction][rate]['v_Vs'] for rate in scan_rates], 
                        y=currents_abs,
                        mode='markers', 
                        name=f'<b>b = {b:.3f}</b>',
                        marker=dict(size=28, color='red', symbol='circle', line=dict(width=4, color='black')),
                        showlegend=True
                    )
                )
                
                # 拟合线
                v_fit = np.linspace(min([cv_data[current_direction][rate]['v_Vs'] for rate in scan_rates]), 
                                    max([cv_data[current_direction][rate]['v_Vs'] for rate in scan_rates]), 100)
                i_fit = np.exp(intercept) * v_fit**b
                fig1.add_trace(
                    go.Scatter(
                        x=v_fit, 
                        y=i_fit,
                        mode='lines', 
                        line=dict(color='blue', width=6, dash='dash'),
                        name='<b>Fitted line</b>',
                        showlegend=True
                    )
                )
                
                fig1 = style_square_axis(fig1, "Scan Rate v (V/s)", "Current |i| (A)", log_x=True, log_y=True)
                fig1.update_layout(
                    title={
                        'text': '<b>i vs v Relationship at 0.1V</b>',
                        'x': 0.5,
                        'xanchor': 'center',
                        'font': {'size': 54, 'color': 'black', 'family': 'Arial Black'}
                    },
                    legend={
                        'font': {'size': 48, 'color': 'black', 'family': 'Arial Black'},
                        'bgcolor': 'rgba(255,255,255,0.95)', 
                        'bordercolor': 'black', 
                        'borderwidth': 3,
                        'x': 0.75, 
                        'y': 0.15
                    }
                )
                
                st.plotly_chart(fig1, use_container_width=True)
                
                # ===== 图4: 贡献率分解堆叠柱状图 (正方形1:1) =====
                fig2 = create_square_figure(1000)
                
                Y = currents_abs / np.sqrt([cv_data[current_direction][rate]['v_Vs'] for rate in scan_rates])
                X = np.sqrt([cv_data[current_direction][rate]['v_Vs'] for rate in scan_rates])
                
                k1, k2, r2, _, _ = stats.linregress(X, Y)
                
                # 计算各扫速下的贡献率
                scan_rates_contrib = []
                for rate in scan_rates:
                    scan_v = rate / 1000
                    sqrt_scan = np.sqrt(scan_v)
                    
                    i_total = k1 * scan_v + k2 * sqrt_scan
                    i_cap = k1 * scan_v
                    cap_ratio_point = (i_cap / i_total) * 100 if i_total > 0 else 50
                    scan_rates_contrib.append(cap_ratio_point)
                
                # 创建堆叠柱状图
                for i, rate in enumerate(scan_rates):
                    fig2.add_trace(go.Bar(
                        name=f'{rate} mV/s',
                        x=[f'{rate} mV/s'],
                        y=[scan_rates_contrib[i]],
                        marker_color='#A569BD',
                        text=[f'<b>{scan_rates_contrib[i]:.1f}%</b>'],
                        textposition='inside',
                        textfont={'size': 36, 'color': 'white', 'family': 'Arial Black'},
                        width=0.5,
                        legendgroup='capacitive',
                        showlegend=False,
                        insidetextanchor='middle'
                    ))
                    
                    fig2.add_trace(go.Bar(
                        name=f'{rate} mV/s',
                        x=[f'{rate} mV/s'],
                        y=[100 - scan_rates_contrib[i]],
                        marker_color='#52BE80',
                        text=[f'<b>{100-scan_rates_contrib[i]:.1f}%</b>'],
                        textposition='inside',
                        textfont={'size': 36, 'color': 'white', 'family': 'Arial Black'},
                        width=0.5,
                        legendgroup='diffusive',
                        showlegend=False,
                        insidetextanchor='middle'
                    ))
                
                fig2 = style_square_axis(fig2, "Scan Rate (mV/s)", "Contribution (%)", y_range=[0, 100])
                fig2.update_layout(
                    title={
                        'text': '<b>Contribution Decomposition at 0.1V</b>',
                        'x': 0.5,
                        'xanchor': 'center',
                        'font': {'size': 54, 'color': 'black', 'family': 'Arial Black'}
                    },
                    barmode='stack',
                    showlegend=False
                )
                
                st.plotly_chart(fig2, use_container_width=True)
                
                # ===== 图5: 区域I和II的b值对比 (使用R² >= 0.99的数据) =====
                fig3 = create_square_figure(1000)
                
                # 重新计算区域I和II的平均b值，只使用R² >= 0.99的数据
                valid_r2_mask_region1 = (b_r2 >= 0.99) & region1_mask
                valid_r2_mask_region2 = (b_r2 >= 0.99) & region2_mask
                
                mean_b1_filtered = np.nanmean(b_values[valid_r2_mask_region1]) if np.any(valid_r2_mask_region1) else np.nan
                mean_b2_filtered = np.nanmean(b_values[valid_r2_mask_region2]) if np.any(valid_r2_mask_region2) else np.nan
                
                # 计算使用的数据点数
                n_points_region1 = np.sum(valid_r2_mask_region1)
                n_points_region2 = np.sum(valid_r2_mask_region2)
                
                fig3.add_trace(
                    go.Bar(
                        x=[f'<b>Region I</b><br><i>Adsorption/Defect</i><br>(n={n_points_region1})', 
                           f'<b>Region II</b><br><i>Insertion/Pore filling</i><br>(n={n_points_region2})'], 
                        y=[mean_b1_filtered, mean_b2_filtered],
                        name='Mean b-value',
                        marker_color=['#5DADE2', '#F1948A'],
                        text=[f'<b>{mean_b1_filtered:.3f}</b>' if not np.isnan(mean_b1_filtered) else '<b>N/A</b>',
                              f'<b>{mean_b2_filtered:.3f}</b>' if not np.isnan(mean_b2_filtered) else '<b>N/A</b>'],
                        textposition='outside',
                        textfont={'size': 48, 'color': 'black', 'family': 'Arial Black'},
                        width=0.5,
                        showlegend=False
                    )
                )
                
                # 添加参考线
                fig3.add_hline(y=0.5, line_dash="dash", line_color="blue", 
                              annotation_text="<b>b=0.5</b>",
                              annotation_position="right",
                              annotation_font=dict(size=42, color='blue', family='Arial Black'))
                fig3.add_hline(y=1.0, line_dash="dash", line_color="green", 
                              annotation_text="<b>b=1.0</b>",
                              annotation_position="right",
                              annotation_font=dict(size=42, color='green', family='Arial Black'))
                
                fig3 = style_square_axis(fig3, "Region", "b-value", y_range=[0, 1.2])
                fig3.update_layout(
                    title={
                        'text': '<b>Comparison of b-values in Two Regions<br>(R² ≥ 0.99)</b>',
                        'x': 0.5,
                        'xanchor': 'center',
                        'font': {'size': 54, 'color': 'black', 'family': 'Arial Black'}
                    }
                )
                
                st.plotly_chart(fig3, use_container_width=True)
                
                # 显示0.1V点分析结果
                st.subheader("0.1V Analysis Results")
                
                result_data = [{
                    'Analysis Item': 'b-value',
                    'Value': f'{b:.3f}',
                    'R²': f'{r_value**2:.3f}',
                    'Mechanism': 'Capacitive dominant' if b > 0.8 else 'Diffusive dominant' if b < 0.6 else 'Mixed control'
                }]
                
                for rate, contrib in zip(scan_rates, scan_rates_contrib):
                    result_data.append({
                        'Analysis Item': f'{rate} mV/s Capacitive Contribution',
                        'Value': f'{contrib:.1f}%',
                        'R²': '-',
                        'Mechanism': '-'
                    })
                
                st.dataframe(pd.DataFrame(result_data), use_container_width=True)
            
            # ==================== Tab 4: 综合报告 ====================
            with tab4:
                st.subheader(f"📋 Comprehensive Report - {direction_label} Scan")
                
                # 计算各区域平均b值（使用R² >= 0.99的数据）
                valid_r2_mask_region1 = (b_r2 >= 0.99) & region1_mask
                valid_r2_mask_region2 = (b_r2 >= 0.99) & region2_mask
                
                mean_b1 = np.nanmean(b_values[valid_r2_mask_region1]) if np.any(valid_r2_mask_region1) else np.nan
                mean_b2 = np.nanmean(b_values[valid_r2_mask_region2]) if np.any(valid_r2_mask_region2) else np.nan
                
                # 计算总的平均电容贡献率
                total_cap_avg = np.nanmean(total_cap_ratios)
                
                # 显示统计信息
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Region I Mean b-value (R²≥0.99)", 
                             f"{mean_b1:.3f}" if not np.isnan(mean_b1) else "N/A")
                    st.info(f"📌 {np.sum(valid_r2_mask_region1)} data points")
                
                with col2:
                    st.metric("Region II Mean b-value (R²≥0.99)", 
                             f"{mean_b2:.3f}" if not np.isnan(mean_b2) else "N/A")
                    st.info(f"📌 {np.sum(valid_r2_mask_region2)} data points")
                
                with col3:
                    st.metric("Average Total Capacitive Contribution", 
                             f"{total_cap_avg:.1f}%" if not np.isnan(total_cap_avg) else "N/A")
                    st.info("📌 Average over 0-0.8V range")
                
                st.markdown("---")
                
                # 机制判断
                st.subheader("🎯 Sodium Storage Mechanism Assessment")
                
                mech_I = "Capacitive dominant (surface adsorption/defects)" if mean_b1 > 0.8 else \
                         "Diffusive dominant" if mean_b1 < 0.6 else "Mixed control"
                mech_II = "Capacitive dominant" if mean_b2 > 0.8 else \
                          "Diffusive dominant (intercalation/pore filling)" if mean_b2 < 0.6 else "Mixed control"
                
                st.markdown(f"""
                - **Region I (Adsorption/Defect, {region1_start}-{region1_end}V)**: {mech_I} (mean b = {mean_b1:.3f}, n={np.sum(valid_r2_mask_region1)})
                - **Region II (Insertion/Pore filling, {region2_start}-{region2_end}V)**: {mech_II} (mean b = {mean_b2:.3f}, n={np.sum(valid_r2_mask_region2)})
                - **At 0.1V boundary**: b = {b:.3f} ({'Capacitive dominant' if b > 0.8 else 'Diffusive dominant' if b < 0.6 else 'Mixed control'})
                """)
                
                # 显示各扫速贡献率汇总
                st.subheader("📊 Contribution Summary at Different Scan Rates")
                contrib_df = pd.DataFrame(scan_contributions)
                st.dataframe(contrib_df, use_container_width=True)
                
                # 下载按钮
                st.markdown("---")
                st.subheader("💾 Download Analysis Report")
                
                # 创建完整的DataFrame用于下载
                download_df = pd.DataFrame({
                    'Potential (V)': v_common,
                    'b-value (corrected)': b_values,
                    'b-value error': b_errors,
                    'b-value R²': b_r2,
                    'k₁ (capacitive coefficient)': k1_values,
                    'k₂ (diffusive coefficient)': k2_values,
                })
                
                # 添加各扫速的贡献率
                for rate in scan_rates:
                    scan_v = rate / 1000
                    sqrt_scan = np.sqrt(scan_v)
                    i_total = k1_values * scan_v + k2_values * sqrt_scan
                    i_cap = k1_values * scan_v
                    cap_ratio = np.where(i_total > 0, (i_cap / i_total) * 100, np.nan)
                    download_df[f'{rate}mV/s Capacitive (%)'] = cap_ratio
                    download_df[f'{rate}mV/s Diffusive (%)'] = 100 - cap_ratio
                
                # 添加区域标记
                download_df['Region'] = 'Other'
                download_df.loc[region1_mask, 'Region'] = 'Region I (Adsorption/Defect)'
                download_df.loc[region2_mask, 'Region'] = 'Region II (Insertion/Pore filling)'
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # 详细数据
                    download_df.to_excel(writer, sheet_name=f'{direction_label}_Detailed', index=False)
                    
                    # 0.1V点分析
                    boundary_df = pd.DataFrame(result_data)
                    boundary_df.to_excel(writer, sheet_name=f'{direction_label}_0.1V_Analysis', index=False)
                    
                    # 各扫速贡献率汇总
                    contrib_df = pd.DataFrame(scan_contributions)
                    contrib_df.to_excel(writer, sheet_name=f'{direction_label}_Contributions', index=False)
                    
                    # 机制说明（使用R²≥0.99的数据）
                    mechanism_df = pd.DataFrame({
                        'Scan Direction': [direction_label, direction_label, direction_label],
                        'Analysis Item': ['Region I (Adsorption/Defect)', 'Region II (Insertion/Pore filling)', '0.1V Boundary'],
                        'Potential Range': [f'{region1_start}-{region1_end} V', 
                                           f'{region2_start}-{region2_end} V',
                                           '0.1V'],
                        'Mean b-value (R²≥0.99)': [f'{mean_b1:.3f}' if not np.isnan(mean_b1) else 'N/A', 
                                                   f'{mean_b2:.3f}' if not np.isnan(mean_b2) else 'N/A',
                                                   f'{b:.3f}'],
                        'Number of points (R²≥0.99)': [np.sum(valid_r2_mask_region1), 
                                                       np.sum(valid_r2_mask_region2),
                                                       '1'],
                        'Dominant Mechanism': [mech_I, mech_II, 
                                              'Capacitive dominant' if b > 0.8 else 'Diffusive dominant' if b < 0.6 else 'Mixed control']
                    })
                    mechanism_df.to_excel(writer, sheet_name=f'{direction_label}_Mechanism', index=False)
                    
                    # 参数说明
                    notes_df = pd.DataFrame({
                        'Parameter': ['Electrode Area A (cm²)', 'Scan Direction',
                                     'Region I Range (V)', 'Region II Range (V)',
                                     'Formula', 'b-value equation', 'Contribution equation',
                                     'R² threshold for averaging'],
                        'Value': [f'{A}', direction_label,
                                  f'{region1_start}-{region1_end}', 
                                  f'{region2_start}-{region2_end}',
                                  r'i = a v^b',
                                  r'log i = b log v + log a',
                                  r'i(V) = k₁v + k₂v^{1/2}',
                                  '0.99']
                    })
                    notes_df.to_excel(writer, sheet_name='Parameters', index=False)
                
                output.seek(0)
                
                st.download_button(
                    label="📥 Download Complete Analysis Report (Excel)",
                    data=output,
                    file_name=f'hard_carbon_cv_analysis_{direction_label}_{pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True
                )
    
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.exception(e)

else:
    # 显示使用说明
    st.info("👆 Please upload Excel file containing five scan rates data")
    
    with st.expander("📝 View Detailed Instructions", expanded=True):
        st.markdown("""
        ### 📋 File Format Requirements
        
        Excel file should contain 10 columns in the following order:
        
        | Column Name | Description |
        |------|------|
        | 0.1_Ewe_V | Voltage data for 0.1 mV/s scan (V) |
        | 0.1_I_mA | Current data for 0.1 mV/s scan (mA) |
        | 0.25_Ewe_V | Voltage data for 0.25 mV/s scan (V) |
        | 0.25_I_mA | Current data for 0.25 mV/s scan (mA) |
        | 0.5_Ewe_V | Voltage data for 0.5 mV/s scan (V) |
        | 0.5_I_mA | Current data for 0.5 mV/s scan (mA) |
        | 1.0_Ewe_V | Voltage data for 1.0 mV/s scan (V) |
        | 1.0_I_mA | Current data for 1.0 mV/s scan (mA) |
        | 2.0_Ewe_V | Voltage data for 2.0 mV/s scan (V) |
        | 2.0_I_mA | Current data for 2.0 mV/s scan (mA) |
        
        ### 🔬 Analysis Methods
        
        #### 1. Forward/Reverse Scan Separation
        The program automatically identifies the turning point of CV curves and separates data into forward (0V→2.5V) and reverse (2.5V→0V) scans.
        
        #### 2. b-value Analysis
        $$i = a v^b$$
        $$\log i = b \log v + \log a$$
        
        - **b = 0.5**: Diffusion-controlled process (bulk insertion/pore filling)
        - **b = 1.0**: Capacitive-controlled process (surface adsorption/defects)
        
        #### 3. Capacitive Contribution Analysis (Dunn method)
        $$i(V) = k_1 v + k_2 v^{1/2}$$
        
        Transformed to:
        $$\frac{i(V)}{v^{1/2}} = k_1 v^{1/2} + k_2$$
        
        #### 4. Two-region Model for Hard Carbon Sodium Storage
        
        - **Region I (Adsorption/Defect, 0.8-0.1V)**: Surface defect site adsorption, capacitive behavior dominant
        - **Region II (Insertion/Pore filling, 0.1-0.01V)**: Interlayer insertion and nanopore filling, diffusive behavior dominant
        
        #### 5. Quality Control
        - Region average b-values are calculated using only data points with R² ≥ 0.99
        - The voltage point with b-value closest to 0.5 in Region II is automatically detected and annotated
        """)