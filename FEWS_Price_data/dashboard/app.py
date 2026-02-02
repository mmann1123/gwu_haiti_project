"""
FEWS NET Haiti Price Dashboard
==============================
Interactive dashboard for visualizing Haiti market price data from FEWS NET.

Run locally:
    streamlit run app.py

Deploy to Streamlit Cloud:
    1. Push to GitHub
    2. Connect repo at share.streamlit.io
    3. Set app path: FEWS_Price_data/dashboard/app.py
"""

import streamlit as st
import duckdb
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from forecasting import fit_all_models, generate_all_forecasts, ForecastResult

# Page config
st.set_page_config(
    page_title="Haiti Food Prices - FEWS NET",
    page_icon="🌾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Database path - works for both local and Streamlit Cloud
DB_PATH = Path(__file__).parent.parent / "database" / "fews_haiti.duckdb"


@st.cache_resource
def get_connection():
    """Get database connection (cached)."""
    return duckdb.connect(str(DB_PATH), read_only=True)


@st.cache_data(ttl=3600)
def get_commodities():
    """Get list of available commodities, with agricultural products first."""
    con = get_connection()
    df = con.execute("""
        SELECT DISTINCT p.name
        FROM products p
        JOIN price_observations po ON p.id = po.product_id
        ORDER BY p.name
    """).fetchdf()

    # Non-agricultural products to put at the bottom
    non_agricultural = {"Charcoal", "Diesel", "Gasoline", "Kerosene"}

    commodities = df["name"].tolist()

    # Sort: agricultural first (alphabetically), then non-agricultural (alphabetically)
    agricultural = sorted([c for c in commodities if c not in non_agricultural])
    fuel_items = sorted([c for c in commodities if c in non_agricultural])

    return agricultural + fuel_items


@st.cache_data(ttl=3600)
def get_markets():
    """Get list of available markets."""
    con = get_connection()
    df = con.execute("""
        SELECT DISTINCT m.name
        FROM markets m
        JOIN price_observations po ON m.id = po.market_id
        ORDER BY m.name
    """).fetchdf()
    return df["name"].tolist()


@st.cache_data(ttl=3600)
def get_mean_prices(commodity: str):
    """Get mean price across all markets for a commodity."""
    con = get_connection()
    df = con.execute("""
        SELECT
            po.period_date,
            AVG(po.value) AS mean_price_htg,
            AVG(po.common_currency_price) AS mean_price_usd,
            MIN(po.value) AS min_price_htg,
            MAX(po.value) AS max_price_htg,
            COUNT(DISTINCT po.market_id) AS num_markets
        FROM price_observations po
        JOIN products p ON po.product_id = p.id
        WHERE p.name = ?
        GROUP BY po.period_date
        ORDER BY po.period_date
    """, [commodity]).fetchdf()
    df["period_date"] = pd.to_datetime(df["period_date"])
    return df


@st.cache_data(ttl=3600)
def get_market_prices(commodity: str):
    """Get individual market prices for a commodity."""
    con = get_connection()
    df = con.execute("""
        SELECT
            m.name AS market,
            po.period_date,
            po.value AS price_htg,
            po.common_currency_price AS price_usd
        FROM price_observations po
        JOIN markets m ON po.market_id = m.id
        JOIN products p ON po.product_id = p.id
        WHERE p.name = ?
        ORDER BY po.period_date, m.name
    """, [commodity]).fetchdf()
    df["period_date"] = pd.to_datetime(df["period_date"])
    return df


@st.cache_data(ttl=3600)
def get_date_range():
    """Get the date range of available data."""
    con = get_connection()
    result = con.execute("""
        SELECT MIN(period_date) AS min_date, MAX(period_date) AS max_date
        FROM price_observations
    """).fetchone()
    return pd.to_datetime(result[0]), pd.to_datetime(result[1])


def calculate_statistics(df: pd.DataFrame, price_col: str) -> dict:
    """Calculate summary statistics for the price data."""
    if df.empty:
        return {}

    latest = df.iloc[-1]
    stats = {
        "current_price": latest[price_col],
        "current_date": latest["period_date"].strftime("%Y-%m-%d"),
    }

    # Month-over-month change
    if len(df) >= 2:
        prev_month = df.iloc[-2][price_col]
        if prev_month and prev_month > 0:
            stats["mom_change"] = ((latest[price_col] - prev_month) / prev_month) * 100

    # Year-over-year change (12 months ago)
    if len(df) >= 13:
        prev_year = df.iloc[-13][price_col]
        if prev_year and prev_year > 0:
            stats["yoy_change"] = ((latest[price_col] - prev_year) / prev_year) * 100

    # 12-month moving average
    if len(df) >= 12:
        stats["moving_avg_12m"] = df[price_col].tail(12).mean()

    return stats


def main():
    # Header
    st.title("🌾 Haiti Food Price Monitor")
    st.markdown("*Data from [FEWS NET](https://fews.net/) - Famine Early Warning Systems Network*")

    # Sidebar controls
    st.sidebar.header("Settings")

    # Commodity selector
    commodities = get_commodities()
    default_idx = commodities.index("Rice (4% Broken)") if "Rice (4% Broken)" in commodities else 0
    selected_commodity = st.sidebar.selectbox(
        "Select Commodity",
        commodities,
        index=default_idx
    )

    # Currency toggle
    currency = st.sidebar.radio(
        "Currency",
        ["HTG (Haitian Gourde)", "USD"],
        index=0
    )
    use_usd = currency == "USD"
    price_col = "mean_price_usd" if use_usd else "mean_price_htg"
    market_price_col = "price_usd" if use_usd else "price_htg"
    currency_symbol = "$" if use_usd else "HTG "

    # Date range
    min_date, max_date = get_date_range()
    date_range = st.sidebar.date_input(
        "Date Range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )

    # Get data
    mean_df = get_mean_prices(selected_commodity)
    market_df = get_market_prices(selected_commodity)

    # Filter by date range
    if len(date_range) == 2:
        start_date, end_date = date_range
        mean_df = mean_df[
            (mean_df["period_date"] >= pd.to_datetime(start_date)) &
            (mean_df["period_date"] <= pd.to_datetime(end_date))
        ]
        market_df = market_df[
            (market_df["period_date"] >= pd.to_datetime(start_date)) &
            (market_df["period_date"] <= pd.to_datetime(end_date))
        ]

    # Statistics sidebar
    st.sidebar.markdown("---")
    st.sidebar.header("Summary Statistics")

    stats = calculate_statistics(mean_df, price_col)
    if stats:
        st.sidebar.metric(
            "Current Price (Mean)",
            f"{currency_symbol}{stats['current_price']:.2f}",
            delta=f"{stats.get('mom_change', 0):.1f}% MoM" if "mom_change" in stats else None
        )

        if "yoy_change" in stats:
            st.sidebar.metric(
                "Year-over-Year Change",
                f"{stats['yoy_change']:.1f}%"
            )

        if "moving_avg_12m" in stats:
            st.sidebar.metric(
                "12-Month Moving Avg",
                f"{currency_symbol}{stats['moving_avg_12m']:.2f}"
            )

        st.sidebar.caption(f"As of {stats['current_date']}")

    # Main content - tabs
    tab1, tab2, tab3 = st.tabs(["📈 Price Trend", "🏪 Market Comparison", "🔮 Price Forecast"])

    with tab1:
        st.subheader(f"Mean Price: {selected_commodity}")

        if mean_df.empty:
            st.warning("No data available for the selected filters.")
        else:
            # Create complete monthly date range and interpolate missing values
            plot_df = mean_df.copy()
            plot_df = plot_df.set_index("period_date")

            # Create complete monthly date range
            full_range = pd.date_range(
                start=plot_df.index.min(),
                end=plot_df.index.max(),
                freq="MS"  # Month start
            )
            # Shift to month end to match data
            full_range = full_range + pd.offsets.MonthEnd(0)

            # Reindex to full range, marking which rows are interpolated
            plot_df = plot_df.reindex(full_range)
            plot_df["is_interpolated"] = plot_df[price_col].isna()

            # Interpolate missing values
            for col in ["mean_price_htg", "mean_price_usd", "min_price_htg", "max_price_htg"]:
                if col in plot_df.columns:
                    plot_df[col] = plot_df[col].interpolate(method="linear")

            plot_df = plot_df.reset_index().rename(columns={"index": "period_date"})

            # Calculate min/max in selected currency
            if use_usd:
                ratio = plot_df["mean_price_usd"] / plot_df["mean_price_htg"]
                min_price = plot_df["min_price_htg"] * ratio
                max_price = plot_df["max_price_htg"] * ratio
            else:
                min_price = plot_df["min_price_htg"]
                max_price = plot_df["max_price_htg"]

            # Create figure
            fig = go.Figure()

            # Add min bound (invisible line for fill reference)
            fig.add_trace(go.Scatter(
                x=plot_df["period_date"],
                y=min_price,
                mode="lines",
                line=dict(width=0),
                showlegend=False,
                hoverinfo="skip"
            ))

            # Add max bound with fill to min
            fig.add_trace(go.Scatter(
                x=plot_df["period_date"],
                y=max_price,
                mode="lines",
                line=dict(width=0),
                fill="tonexty",
                fillcolor="rgba(31, 119, 180, 0.15)",
                name="Price Range (Min-Max)",
                hoverinfo="skip"
            ))

            # Split data into actual and interpolated segments for different colors
            # Add actual data points (blue)
            actual_mask = ~plot_df["is_interpolated"]
            fig.add_trace(go.Scatter(
                x=plot_df.loc[actual_mask, "period_date"],
                y=plot_df.loc[actual_mask, price_col],
                mode="lines+markers",
                name="Actual Price",
                line=dict(color="#1f77b4", width=2),
                marker=dict(size=4),
                hovertemplate=f"Date: %{{x|%Y-%m}}<br>Price: {currency_symbol}%{{y:.2f}}<extra></extra>"
            ))

            # Add interpolated segments (red dashed)
            # Find runs of interpolated points and connect them to actual points
            interp_mask = plot_df["is_interpolated"]
            if interp_mask.any():
                # Create segments that include interpolated points and their neighbors
                plot_df["segment"] = (~interp_mask).cumsum()
                for seg_id in plot_df.loc[interp_mask, "segment"].unique():
                    # Get interpolated points in this segment
                    seg_mask = (plot_df["segment"] == seg_id) & interp_mask
                    seg_indices = plot_df[seg_mask].index.tolist()

                    if seg_indices:
                        # Include one point before and after for continuity
                        start_idx = max(0, seg_indices[0] - 1)
                        end_idx = min(len(plot_df) - 1, seg_indices[-1] + 1)
                        seg_data = plot_df.iloc[start_idx:end_idx + 1]

                        fig.add_trace(go.Scatter(
                            x=seg_data["period_date"],
                            y=seg_data[price_col],
                            mode="lines",
                            line=dict(color="red", width=2, dash="dot"),
                            showlegend=False,
                            hovertemplate=f"Date: %{{x|%Y-%m}}<br>Price: {currency_symbol}%{{y:.2f}} (interpolated)<extra></extra>"
                        ))

                # Add legend entry for interpolated
                fig.add_trace(go.Scatter(
                    x=[None], y=[None],
                    mode="lines",
                    line=dict(color="red", width=2, dash="dot"),
                    name="Interpolated (missing data)"
                ))

            fig.update_layout(
                xaxis_title="Date",
                yaxis_title=f"Price ({currency.split()[0]})",
                hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
                margin=dict(l=0, r=0, t=30, b=0)
            )

            st.plotly_chart(fig, use_container_width=True)

            # Show data table
            with st.expander("View Data"):
                display_df = mean_df[["period_date", price_col, "num_markets"]].copy()
                display_df.columns = ["Date", f"Mean Price ({currency.split()[0]})", "# Markets"]
                display_df["Date"] = display_df["Date"].dt.strftime("%Y-%m")
                st.dataframe(display_df.tail(24), use_container_width=True)

    with tab2:
        st.subheader(f"Market Comparison: {selected_commodity}")

        # Market selector
        markets = get_markets()
        selected_markets = st.multiselect(
            "Select Markets to Compare",
            markets,
            default=markets[:5]  # Default to first 5 markets
        )

        if market_df.empty:
            st.warning("No data available for the selected filters.")
        elif not selected_markets:
            st.info("Select one or more markets to compare.")
        else:
            # Filter to selected markets
            filtered_market_df = market_df[market_df["market"].isin(selected_markets)]

            # Pivot for easier plotting
            pivot_df = filtered_market_df.pivot(
                index="period_date",
                columns="market",
                values=market_price_col
            ).reset_index()

            # Create figure
            fig = go.Figure()

            # Add mean line (bold)
            fig.add_trace(go.Scatter(
                x=mean_df["period_date"],
                y=mean_df[price_col],
                mode="lines",
                name="Mean (All Markets)",
                line=dict(color="black", width=3),
                hovertemplate=f"Mean: {currency_symbol}%{{y:.2f}}<extra></extra>"
            ))

            # Add individual market lines
            colors = px.colors.qualitative.Set2
            for i, market in enumerate(selected_markets):
                if market in pivot_df.columns:
                    fig.add_trace(go.Scatter(
                        x=pivot_df["period_date"],
                        y=pivot_df[market],
                        mode="lines",
                        name=market,
                        line=dict(color=colors[i % len(colors)], width=1.5),
                        opacity=0.7,
                        hovertemplate=f"{market}: {currency_symbol}%{{y:.2f}}<extra></extra>"
                    ))

            fig.update_layout(
                xaxis_title="Date",
                yaxis_title=f"Price ({currency.split()[0]})",
                hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
                margin=dict(l=0, r=0, t=30, b=0)
            )

            st.plotly_chart(fig, use_container_width=True)

            # Show latest prices table
            with st.expander("Latest Prices by Market"):
                latest_date = filtered_market_df["period_date"].max()
                latest_df = filtered_market_df[filtered_market_df["period_date"] == latest_date][
                    ["market", market_price_col]
                ].copy()
                latest_df.columns = ["Market", f"Price ({currency.split()[0]})"]
                latest_df = latest_df.sort_values(f"Price ({currency.split()[0]})", ascending=False)
                st.dataframe(latest_df, use_container_width=True)

    with tab3:
        st.subheader(f"8-Month Price Forecast: {selected_commodity}")
        st.markdown("*Forecasts generated using Facebook Prophet with automatic seasonality detection*")
        
        # Initialize session state for model caching
        if 'forecast_models' not in st.session_state:
            st.session_state.forecast_models = {}
        if 'forecast_product' not in st.session_state:
            st.session_state.forecast_product = None
        if 'forecast_currency' not in st.session_state:
            st.session_state.forecast_currency = None
        
        # Check if we need to refresh models
        need_refresh = (
            st.session_state.forecast_product != selected_commodity or
            st.session_state.forecast_currency != currency
        )
        
        # Controls
        col1, col2 = st.columns([3, 1])
        with col1:
            forecast_horizon = st.slider(
                "Forecast Horizon (Months)",
                min_value=1,
                max_value=8,
                value=8,
                help="Number of months to forecast into the future"
            )
        with col2:
            refresh_button = st.button("🔄 Refresh Models", help="Re-train Prophet models with latest data")
        
        if refresh_button:
            need_refresh = True
            st.session_state.forecast_models = {}
        
        # Fit models if needed
        if need_refresh or not st.session_state.forecast_models:
            with st.spinner("Training Prophet models... This may take a minute."):
                results, availability = fit_all_models(
                    str(DB_PATH),
                    selected_commodity,
                    currency='USD' if use_usd else 'HTG',
                    min_months=24
                )
                
                st.session_state.forecast_models = results
                st.session_state.forecast_availability = availability
                st.session_state.forecast_product = selected_commodity
                st.session_state.forecast_currency = currency
        else:
            results = st.session_state.forecast_models
            availability = st.session_state.forecast_availability
        
        if not results:
            st.warning("No data available for forecasting. This commodity may not have sufficient historical data.")
        else:
            # Generate forecasts
            forecasts = generate_all_forecasts(results, periods=forecast_horizon)
            
            # Market selector with availability info
            available_markets = [m for m in results.keys() if results[m].success and m != 'Market Average']
            
            # Create market options with tooltips
            market_options = []
            disabled_markets = []
            
            for market_name, info in availability.items():
                if info['sufficient']:
                    market_options.append(market_name)
                else:
                    disabled_markets.append(market_name)
            
            # Show data availability info
            with st.expander("📊 Data Availability"):
                avail_data = []
                for market_name, info in availability.items():
                    status = "✅ Included" if info['sufficient'] else "❌ Excluded"
                    reason = info['reason'] if info['reason'] else "Sufficient data"
                    avail_data.append({
                        'Market': market_name,
                        'Observations': info['n_observations'],
                        'Time Span (months)': info['months_span'],
                        'Status': status,
                        'Notes': reason
                    })
                avail_df = pd.DataFrame(avail_data)
                st.dataframe(avail_df, use_container_width=True)
                
                if disabled_markets:
                    st.info(f"**Excluded markets:** {', '.join(disabled_markets)} (insufficient data: require 24+ months)")
            
            # View selector
            view_mode = st.radio(
                "View Mode",
                ["Market Average", "Individual Markets"],
                horizontal=True
            )
            
            if view_mode == "Market Average":
                # Show market average forecast
                if 'Market Average' in forecasts:
                    forecast_df = forecasts['Market Average']
                    model = results['Market Average'].model
                    
                    # Split historical and future
                    historical_df = forecast_df[forecast_df['ds'] <= forecast_df['ds'].max() - pd.DateOffset(months=forecast_horizon)]
                    future_df = forecast_df[forecast_df['ds'] > forecast_df['ds'].max() - pd.DateOffset(months=forecast_horizon)]
                    
                    # Create figure
                    fig = go.Figure()
                    
                    # Historical actuals
                    fig.add_trace(go.Scatter(
                        x=historical_df['ds'].tail(36),  # Show last 36 months
                        y=historical_df['yhat'].tail(36),
                        mode='lines',
                        name='Historical (Fitted)',
                        line=dict(color='blue', width=2),
                        hovertemplate=f"Date: %{{x}}<br>Price: {currency_symbol}%{{y:.2f}}<extra></extra>"
                    ))
                    
                    # Forecast
                    fig.add_trace(go.Scatter(
                        x=future_df['ds'],
                        y=future_df['yhat'],
                        mode='lines',
                        name='Forecast',
                        line=dict(color='red', width=2, dash='dash'),
                        hovertemplate=f"Date: %{{x}}<br>Forecast: {currency_symbol}%{{y:.2f}}<extra></extra>"
                    ))
                    
                    # 95% confidence interval
                    fig.add_trace(go.Scatter(
                        x=future_df['ds'].tolist() + future_df['ds'].tolist()[::-1],
                        y=future_df['yhat_upper'].tolist() + future_df['yhat_lower'].tolist()[::-1],
                        fill='toself',
                        fillcolor='rgba(255, 0, 0, 0.1)',
                        line=dict(color='rgba(255, 0, 0, 0)'),
                        name='95% Confidence',
                        hoverinfo='skip',
                        showlegend=True
                    ))
                    
                    # 80% confidence interval
                    fig.add_trace(go.Scatter(
                        x=future_df['ds'].tolist() + future_df['ds'].tolist()[::-1],
                        y=(future_df['yhat'] + (future_df['yhat_upper'] - future_df['yhat']) * 0.8).tolist() + 
                          (future_df['yhat'] - (future_df['yhat'] - future_df['yhat_lower']) * 0.8).tolist()[::-1],
                        fill='toself',
                        fillcolor='rgba(255, 0, 0, 0.2)',
                        line=dict(color='rgba(255, 0, 0, 0)'),
                        name='80% Confidence',
                        hoverinfo='skip',
                        showlegend=True
                    ))
                    
                    fig.update_layout(
                        title=f"Market Average Forecast - {selected_commodity}",
                        xaxis_title="Date",
                        yaxis_title=f"Price ({currency.split()[0]})",
                        hovermode="x unified",
                        legend=dict(orientation="h", yanchor="bottom", y=1.02),
                        height=500
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Forecast table
                    with st.expander("📋 Forecast Table"):
                        table_df = future_df[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].copy()
                        table_df.columns = ['Date', 'Forecast', '95% Lower', '95% Upper']
                        table_df['Date'] = table_df['Date'].dt.strftime('%Y-%m')
                        for col in ['Forecast', '95% Lower', '95% Upper']:
                            table_df[col] = table_df[col].apply(lambda x: f"{currency_symbol}{x:.2f}")
                        st.dataframe(table_df, use_container_width=True)
                    
                    # Prophet components
                    with st.expander("🔍 Model Components (Trend & Seasonality)"):
                        from prophet.plot import plot_components_plotly
                        components_fig = plot_components_plotly(model, forecast_df)
                        st.plotly_chart(components_fig, use_container_width=True)
                        st.caption("Prophet automatically detects and separates trend and seasonal patterns")
                
                else:
                    st.error("Market average forecast failed to generate.")
                    if 'Market Average' in results and results['Market Average'].error:
                        st.error(f"Error: {results['Market Average'].error}")
            
            else:  # Individual Markets
                selected_forecast_markets = st.multiselect(
                    "Select Markets to Display",
                    market_options,
                    default=market_options[:3] if len(market_options) >= 3 else market_options,
                    help="Choose which markets to show in the forecast"
                )
                
                if not selected_forecast_markets:
                    st.info("Select at least one market to view forecasts.")
                else:
                    # Create combined plot
                    fig = go.Figure()
                    colors = px.colors.qualitative.Set2
                    
                    for i, market_name in enumerate(selected_forecast_markets):
                        if market_name in forecasts:
                            forecast_df = forecasts[market_name]
                            
                            # Split historical and future
                            historical_df = forecast_df[forecast_df['ds'] <= forecast_df['ds'].max() - pd.DateOffset(months=forecast_horizon)]
                            future_df = forecast_df[forecast_df['ds'] > forecast_df['ds'].max() - pd.DateOffset(months=forecast_horizon)]
                            
                            color = colors[i % len(colors)]
                            
                            # Historical
                            fig.add_trace(go.Scatter(
                                x=historical_df['ds'].tail(36),
                                y=historical_df['yhat'].tail(36),
                                mode='lines',
                                name=f'{market_name}',
                                line=dict(color=color, width=2),
                                legendgroup=market_name,
                                hovertemplate=f"{market_name}<br>Date: %{{x}}<br>Price: {currency_symbol}%{{y:.2f}}<extra></extra>"
                            ))
                            
                            # Forecast
                            fig.add_trace(go.Scatter(
                                x=future_df['ds'],
                                y=future_df['yhat'],
                                mode='lines',
                                name=f'{market_name} (Forecast)',
                                line=dict(color=color, width=2, dash='dash'),
                                legendgroup=market_name,
                                showlegend=False,
                                hovertemplate=f"{market_name} Forecast<br>Date: %{{x}}<br>Price: {currency_symbol}%{{y:.2f}}<extra></extra>"
                            ))
                            
                            # Confidence interval (lighter)
                            fig.add_trace(go.Scatter(
                                x=future_df['ds'].tolist() + future_df['ds'].tolist()[::-1],
                                y=future_df['yhat_upper'].tolist() + future_df['yhat_lower'].tolist()[::-1],
                                fill='toself',
                                fillcolor=f'rgba{tuple(list(px.colors.hex_to_rgb(color)) + [0.1])}',
                                line=dict(color='rgba(255,255,255,0)'),
                                showlegend=False,
                                legendgroup=market_name,
                                hoverinfo='skip'
                            ))
                        else:
                            # Show error for this market
                            if market_name in results and not results[market_name].success:
                                st.error(f"**{market_name}**: {results[market_name].error}")
                    
                    fig.update_layout(
                        title=f"Individual Market Forecasts - {selected_commodity}",
                        xaxis_title="Date",
                        yaxis_title=f"Price ({currency.split()[0]})",
                        hovermode="x unified",
                        legend=dict(orientation="v", yanchor="top", y=1),
                        height=600
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Forecast comparison table
                    with st.expander("📋 Forecast Comparison Table"):
                        comparison_data = []
                        for market_name in selected_forecast_markets:
                            if market_name in forecasts:
                                forecast_df = forecasts[market_name]
                                future_df = forecast_df[forecast_df['ds'] > forecast_df['ds'].max() - pd.DateOffset(months=forecast_horizon)]
                                
                                for _, row in future_df.iterrows():
                                    comparison_data.append({
                                        'Market': market_name,
                                        'Date': row['ds'].strftime('%Y-%m'),
                                        'Forecast': f"{currency_symbol}{row['yhat']:.2f}",
                                        '95% Lower': f"{currency_symbol}{row['yhat_lower']:.2f}",
                                        '95% Upper': f"{currency_symbol}{row['yhat_upper']:.2f}"
                                    })
                        
                        if comparison_data:
                            comparison_df = pd.DataFrame(comparison_data)
                            st.dataframe(comparison_df, use_container_width=True)

    # Footer
    st.markdown("---")
    st.caption(
        "Data source: [FEWS NET Data Warehouse](https://fdw.fews.net/) | "
        "Price data collected by CNSA/FEWS NET Haiti | "
        "Updated monthly"
    )


if __name__ == "__main__":
    main()
