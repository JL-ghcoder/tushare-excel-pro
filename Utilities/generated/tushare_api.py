import requests
import json
import pandas as pd

def call_tushare_api(api_name, token, **params):
    """通用Tushare API调用函数
    
    Args:
        api_name (str): API名称
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    url = "http://api.tushare.pro"
    
    # 构建请求数据
    payload = {
        "api_name": api_name,
        "token": token,
        "params": params,
        "fields": ""
    }
    
    # 发送请求
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json"
    }
    
    try:
        response = requests.post(url, data=json.dumps(payload), headers=headers)
        response.raise_for_status()  # 检查HTTP错误
        
        # 解析响应
        result = response.json()
        
        if result.get('code') == 0:
            # 成功获取数据
            data = result.get('data')
            if data and 'items' in data and 'fields' in data:
                # 转换为DataFrame
                df = pd.DataFrame(data['items'], columns=data['fields'])
                return df
            else:
                raise ValueError("API返回的数据格式不正确")
        else:
            # API返回错误
            raise ValueError(f"API错误: {result.get('msg')}")
    
    except Exception as e:
        raise Exception(f"调用Tushare API时出错: {str(e)}")


def tushare_fund_basic(token, **params):
    """
    基金列表
    
    获取公募基金数据列表，包括场内和场外基金
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fund_basic", token, **params)

def tushare_stock_basic(token, **params):
    """
    股票列表
    
    获取基础信息数据，包括股票代码、名称、上市日期、退市日期等
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stock_basic", token, **params)

def tushare_trade_cal(token, **params):
    """
    交易日历
    
    获取各大期货交易所交易日历数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("trade_cal", token, **params)

def tushare_daily(token, **params):
    """
    日线行情
    
    获取股票行情数据，或通过[**通用行情接口**]( https://tushare.pro/document/2?doc_id=109)获取数据，包含了前后复权数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("daily", token, **params)

def tushare_adj_factor(token, **params):
    """
    复权因子
    
    获取股票复权因子，可提取单只股票全部历史复权因子，也可以提取单日全部股票的复权因子。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("adj_factor", token, **params)

def tushare_daily_basic(token, **params):
    """
    每日指标
    
    获取全部股票每日重要的基本面指标，可用于选股分析、报表展示等。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("daily_basic", token, **params)

def tushare_income(token, **params):
    """
    利润表
    
    获取上市公司财务利润表数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("income", token, **params)

def tushare_balancesheet(token, **params):
    """
    资产负债表
    
    获取上市公司资产负债表
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("balancesheet", token, **params)

def tushare_cashflow(token, **params):
    """
    现金流量表
    
    获取上市公司现金流量表
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cashflow", token, **params)

def tushare_forecast(token, **params):
    """
    业绩预告
    
    获取业绩预告数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("forecast", token, **params)

def tushare_express(token, **params):
    """
    业绩快报
    
    获取上市公司业绩快报
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("express", token, **params)

def tushare_moneyflow_hsgt(token, **params):
    """
    沪深港通资金流向
    
    获取沪股通、深股通、港股通每日资金流向数据，每次最多返回300条记录，总量不限制。每天18~20点之间完成当日更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("moneyflow_hsgt", token, **params)

def tushare_hsgt_top10(token, **params):
    """
    沪深股通十大成交股
    
    获取沪股通、深股通每日前十大成交详细数据，每天18~20点之间完成当日更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("hsgt_top10", token, **params)

def tushare_ggt_top10(token, **params):
    """
    港股通十大成交股
    
    获取港股通每日成交数据，其中包括沪市、深市详细数据，每天18~20点之间完成当日更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("ggt_top10", token, **params)

def tushare_margin(token, **params):
    """
    融资融券交易汇总
    
    获取融资融券每日交易汇总数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("margin", token, **params)

def tushare_margin_detail(token, **params):
    """
    融资融券交易明细
    
    获取沪深两市每日融资融券明细
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("margin_detail", token, **params)

def tushare_top10_holders(token, **params):
    """
    前十大股东
    
    获取上市公司前十大股东数据，包括持有数量和比例等信息
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("top10_holders", token, **params)

def tushare_top10_floatholders(token, **params):
    """
    前十大流通股东
    
    获取上市公司前十大流通股东数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("top10_floatholders", token, **params)

def tushare_fina_indicator(token, **params):
    """
    财务指标数据
    
    获取上市公司财务指标数据，为避免服务器压力，现阶段每次请求最多返回60条记录，可通过设置日期多次请求获取更多数据。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fina_indicator", token, **params)

def tushare_fina_audit(token, **params):
    """
    财务审计意见
    
    获取上市公司定期财务审计意见数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fina_audit", token, **params)

def tushare_fina_mainbz(token, **params):
    """
    主营业务构成
    
    获得上市公司主营业务构成，分地区和产品两种方式
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fina_mainbz", token, **params)

def tushare_tmt_twincomedetail(token, **params):
    """
    台湾电子产业月营收明细
    
    获取台湾TMT行业上市公司各类产品月度营收情况。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("tmt_twincomedetail", token, **params)

def tushare_tmt_twincome(token, **params):
    """
    台湾电子产业月营收
    
    获取台湾TMT电子产业领域各类产品月度营收数据。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("tmt_twincome", token, **params)

def tushare_index_basic(token, **params):
    """
    指数基本信息
    
    获取指数基础信息。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("index_basic", token, **params)

def tushare_index_daily(token, **params):
    """
    南华期货指数行情
    
    获取南华指数每日行情，指数行情也可以通过[**通用行情接口**]( https://tushare.pro/document/2?doc_id=109)获取数据．
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("index_daily", token, **params)

def tushare_index_weight(token, **params):
    """
    指数成分和权重
    
    获取各类指数成分和权重，**月度数据** ，建议输入参数里开始日期和结束日分别输入当月第一天和最后一天的日期。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("index_weight", token, **params)

def tushare_namechange(token, **params):
    """
    股票曾用名
    
    历史名称变更记录
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("namechange", token, **params)

def tushare_dividend(token, **params):
    """
    分红送股数据
    
    分红送股数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("dividend", token, **params)

def tushare_hs_const(token, **params):
    """
    沪深股通成分股
    
    获取沪股通、深股通成分数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("hs_const", token, **params)

def tushare_top_list(token, **params):
    """
    龙虎榜每日统计单
    
    龙虎榜每日交易明细
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("top_list", token, **params)

def tushare_top_inst(token, **params):
    """
    龙虎榜机构交易单
    
    龙虎榜机构成交明细
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("top_inst", token, **params)

def tushare_pledge_stat(token, **params):
    """
    股权质押统计数据
    
    获取股票质押统计数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("pledge_stat", token, **params)

def tushare_pledge_detail(token, **params):
    """
    股权质押明细数据
    
    获取股票质押明细数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("pledge_detail", token, **params)

def tushare_stock_company(token, **params):
    """
    上市公司基本信息
    
    获取上市公司基础信息，单次提取4500条，可以根据交易所分批提取
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stock_company", token, **params)

def tushare_bo_monthly(token, **params):
    """
    电影月度票房
    
    获取电影月度票房数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("bo_monthly", token, **params)

def tushare_bo_weekly(token, **params):
    """
    电影周度票房
    
    获取周度票房数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("bo_weekly", token, **params)

def tushare_bo_daily(token, **params):
    """
    电影日度票房
    
    获取电影日度票房
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("bo_daily", token, **params)

def tushare_bo_cinema(token, **params):
    """
    影院日度票房
    
    获取每日各影院的票房数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("bo_cinema", token, **params)

def tushare_fund_company(token, **params):
    """
    基金管理人
    
    获取公募基金管理人列表
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fund_company", token, **params)

def tushare_fund_nav(token, **params):
    """
    基金净值
    
    获取公募基金净值数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fund_nav", token, **params)

def tushare_fund_div(token, **params):
    """
    基金分红
    
    获取公募基金分红数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fund_div", token, **params)

def tushare_fund_portfolio(token, **params):
    """
    基金持仓
    
    获取公募基金持仓数据，季度更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fund_portfolio", token, **params)

def tushare_new_share(token, **params):
    """
    IPO新股上市
    
    获取新股上市列表数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("new_share", token, **params)

def tushare_repurchase(token, **params):
    """
    股票回购
    
    获取上市公司回购股票数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("repurchase", token, **params)

def tushare_concept(token, **params):
    """
    概念股分类表
    
    获取概念股分类，目前只有ts一个来源，未来将逐步增加来源
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("concept", token, **params)

def tushare_concept_detail(token, **params):
    """
    概念股明细列表
    
    获取概念股分类明细数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("concept_detail", token, **params)

def tushare_fund_daily(token, **params):
    """
    基金行情（含ETF）
    
    获取场内基金日线行情，类似股票日行情，包括ETF行情
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fund_daily", token, **params)

def tushare_index_dailybasic(token, **params):
    """
    大盘指数每日指标
    
    目前只提供上证综指，深证成指，上证50，中证500，中小板指，创业板指的每日指标数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("index_dailybasic", token, **params)

def tushare_fut_basic(token, **params):
    """
    合约信息
    
    获取期货合约列表数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fut_basic", token, **params)

def tushare_fut_daily(token, **params):
    """
    日线行情
    
    期货日线行情数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fut_daily", token, **params)

def tushare_fut_holding(token, **params):
    """
    每日持仓排名
    
    获取每日成交持仓排名数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fut_holding", token, **params)

def tushare_fut_wsr(token, **params):
    """
    仓单日报
    
    获取仓单日报数据，了解各仓库/厂库的仓单变化
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fut_wsr", token, **params)

def tushare_fut_settle(token, **params):
    """
    每日结算参数
    
    获取每日结算参数数据，包括交易和交割费率等
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fut_settle", token, **params)

def tushare_news(token, **params):
    """
    新闻快讯
    
    获取主流新闻网站的快讯新闻数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("news", token, **params)

def tushare_weekly(token, **params):
    """
    周线行情
    
    获取A股周线行情
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("weekly", token, **params)

def tushare_monthly(token, **params):
    """
    月线行情
    
    获取A股月线数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("monthly", token, **params)

def tushare_shibor(token, **params):
    """
    Shibor利率
    
    shibor利率
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("shibor", token, **params)

def tushare_shibor_quote(token, **params):
    """
    Shibor报价数据
    
    Shibor报价数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("shibor_quote", token, **params)

def tushare_shibor_lpr(token, **params):
    """
    LPR贷款基础利率
    
    LPR贷款基础利率
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("shibor_lpr", token, **params)

def tushare_libor(token, **params):
    """
    Libor利率
    
    Libor拆借利率
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("libor", token, **params)

def tushare_hibor(token, **params):
    """
    Hibor利率
    
    Hibor利率
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("hibor", token, **params)

def tushare_cctv_news(token, **params):
    """
    新闻联播文字稿
    
    获取新闻联播文字稿数据，数据开始于2006年6月，超过12年历史
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cctv_news", token, **params)

def tushare_film_record(token, **params):
    """
    全国电影剧本备案数据
    
    获取全国电影剧本备案的公示数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("film_record", token, **params)

def tushare_opt_basic(token, **params):
    """
    期权合约信息
    
    获取期权合约信息
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("opt_basic", token, **params)

def tushare_opt_daily(token, **params):
    """
    期权日线行情
    
    获取期权日线行情
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("opt_daily", token, **params)

def tushare_share_float(token, **params):
    """
    限售股解禁
    
    获取限售股解禁
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("share_float", token, **params)

def tushare_block_trade(token, **params):
    """
    大宗交易
    
    大宗交易
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("block_trade", token, **params)

def tushare_disclosure_date(token, **params):
    """
    财报披露日期表
    
    获取财报披露计划日期
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("disclosure_date", token, **params)

def tushare_stk_account(token, **params):
    """
    股票开户数据（停）
    
    获取股票账户开户数据，统计周期为一周
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_account", token, **params)

def tushare_stk_account_old(token, **params):
    """
    股票开户数据（旧）
    
    获取股票账户开户数据旧版格式数据，数据从2008年1月开始，到2015年5月29，新数据请通过[股票开户数据](https://tushare.pro/document/2?doc_id=164)获取。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_account_old", token, **params)

def tushare_stk_holdernumber(token, **params):
    """
    股东人数
    
    获取上市公司股东户数数据，数据不定期公布
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_holdernumber", token, **params)

def tushare_moneyflow(token, **params):
    """
    个股资金流向
    
    获取沪深A股票资金流向数据，分析大单小单成交情况，用于判别资金动向，数据开始于2010年。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("moneyflow", token, **params)

def tushare_index_weekly(token, **params):
    """
    指数周线行情
    
    获取指数周线行情
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("index_weekly", token, **params)

def tushare_index_monthly(token, **params):
    """
    指数月线行情
    
    获取指数月线行情,每月更新一次
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("index_monthly", token, **params)

def tushare_wz_index(token, **params):
    """
    温州民间借贷利率
    
    温州民间借贷利率，即温州指数
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("wz_index", token, **params)

def tushare_gz_index(token, **params):
    """
    广州民间借贷利率
    
    广州民间借贷利率
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("gz_index", token, **params)

def tushare_stk_holdertrade(token, **params):
    """
    股东增减持
    
    获取上市公司增减持数据，了解重要股东近期及历史上的股份增减变化
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_holdertrade", token, **params)

def tushare_anns_d(token, **params):
    """
    上市公司公告
    
    获取全量公告数据，提供pdf下载URL
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("anns_d", token, **params)

def tushare_fx_obasic(token, **params):
    """
    外汇基础信息（海外）
    
    获取海外外汇基础信息，目前只有FXCM交易商的数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fx_obasic", token, **params)

def tushare_fx_daily(token, **params):
    """
    外汇日线行情
    
    获取外汇日线行情
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fx_daily", token, **params)

def tushare_teleplay_record(token, **params):
    """
    全国电视剧备案公示数据
    
    获取2009年以来全国拍摄制作电视剧备案公示数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("teleplay_record", token, **params)

def tushare_index_classify(token, **params):
    """
    申万行业分类
    
    获取申万行业分类，可以获取申万2014年版本（28个一级分类，104个二级分类，227个三级分类）和2021年本版（31个一级分类，134个二级分类，346个三级分类）列表信息
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("index_classify", token, **params)

def tushare_stk_limit(token, **params):
    """
    每日涨跌停价格
    
    获取全市场（包含A/B股和基金）每日涨跌停价格，包括涨停价格，跌停价格等，每个交易日8点40左右更新当日股票涨跌停价格。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_limit", token, **params)

def tushare_cb_basic(token, **params):
    """
    可转债基础信息
    
    获取可转债基本信息
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cb_basic", token, **params)

def tushare_cb_issue(token, **params):
    """
    可转债发行
    
    获取可转债发行数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cb_issue", token, **params)

def tushare_cb_daily(token, **params):
    """
    可转债行情
    
    获取可转债行情
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cb_daily", token, **params)

def tushare_hk_hold(token, **params):
    """
    沪深股通持股明细
    
    获取沪深港股通持股明细，数据来源港交所。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("hk_hold", token, **params)

def tushare_fut_mapping(token, **params):
    """
    期货主力与连续合约
    
    获取期货主力（或连续）合约与月合约映射数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fut_mapping", token, **params)

def tushare_hk_basic(token, **params):
    """
    港股基础信息
    
    获取港股列表信息
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("hk_basic", token, **params)

def tushare_hk_daily(token, **params):
    """
    港股日线行情
    
    获取港股每日增量和历史行情，每日18点左右更新当日数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("hk_daily", token, **params)

def tushare_stk_managers(token, **params):
    """
    上市公司管理层
    
    获取上市公司管理层
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_managers", token, **params)

def tushare_stk_rewards(token, **params):
    """
    管理层薪酬和持股
    
    获取上市公司管理层薪酬和持股
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_rewards", token, **params)

def tushare_major_news(token, **params):
    """
    新闻通讯（长篇）
    
    获取长篇通讯信息，覆盖主要新闻资讯网站
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("major_news", token, **params)

def tushare_ggt_daily(token, **params):
    """
    港股通每日成交统计
    
    获取港股通每日成交信息，数据从2014年开始
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("ggt_daily", token, **params)

def tushare_ggt_monthly(token, **params):
    """
    港股通每月成交统计
    
    港股通每月成交信息，数据从2014年开始
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("ggt_monthly", token, **params)

def tushare_fund_adj(token, **params):
    """
    复权因子
    
    获取基金复权因子，用于计算基金复权行情
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fund_adj", token, **params)

def tushare_yc_cb(token, **params):
    """
    国债收益率曲线
    
    获取中债收益率曲线，目前可获取中债国债收益率曲线即期和到期收益率曲线数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("yc_cb", token, **params)

def tushare_fund_share(token, **params):
    """
    基金规模
    
    获取基金规模数据，包含上海和深圳ETF基金
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fund_share", token, **params)

def tushare_fund_manager(token, **params):
    """
    基金经理
    
    获取公募基金经理数据，包括基金经理简历等数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fund_manager", token, **params)

def tushare_index_global(token, **params):
    """
    国际主要指数
    
    获取国际主要指数日线行情
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("index_global", token, **params)

def tushare_suspend_d(token, **params):
    """
    每日停复牌信息
    
    按日期方式获取股票每日停复牌信息
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("suspend_d", token, **params)

def tushare_daily_info(token, **params):
    """
    沪深市场每日交易统计
    
    获取交易所股票交易统计，包括各板块明细
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("daily_info", token, **params)

def tushare_fut_weekly_detail(token, **params):
    """
    期货主要品种交易周报
    
    获取期货交易所主要品种每周交易统计信息，数据从2010年3月开始
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fut_weekly_detail", token, **params)

def tushare_us_tycr(token, **params):
    """
    国债收益率曲线利率
    
    获取美国每日国债收益率曲线利率
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("us_tycr", token, **params)

def tushare_us_trycr(token, **params):
    """
    国债实际收益率曲线利率
    
    国债实际收益率曲线利率
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("us_trycr", token, **params)

def tushare_us_tbr(token, **params):
    """
    短期国债利率
    
    获取美国短期国债利率数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("us_tbr", token, **params)

def tushare_us_tltr(token, **params):
    """
    国债长期利率
    
    国债长期利率
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("us_tltr", token, **params)

def tushare_us_trltr(token, **params):
    """
    国债长期利率平均值
    
    国债实际长期利率平均值
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("us_trltr", token, **params)

def tushare_cn_gdp(token, **params):
    """
    国内生产总值（GDP）
    
    获取国民经济之GDP数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cn_gdp", token, **params)

def tushare_cn_cpi(token, **params):
    """
    居民消费价格指数（CPI）
    
    获取CPI居民消费价格数据，包括全国、城市和农村的数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cn_cpi", token, **params)

def tushare_eco_cal(token, **params):
    """
    全球财经事件
    
    获取全球财经日历、包括经济事件数据更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("eco_cal", token, **params)

def tushare_cn_m(token, **params):
    """
    货币供应量（月）
    
    获取货币供应量之月度数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cn_m", token, **params)

def tushare_cn_ppi(token, **params):
    """
    工业生产者出厂价格指数（PPI）
    
    获取PPI工业生产者出厂价格指数数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cn_ppi", token, **params)

def tushare_cb_price_chg(token, **params):
    """
    可转债转股价变动
    
    获取可转债转股价变动
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cb_price_chg", token, **params)

def tushare_cb_share(token, **params):
    """
    可转债转股结果
    
    获取可转债转股结果
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cb_share", token, **params)

def tushare_hk_tradecal(token, **params):
    """
    港股交易日历
    
    获取交易日历
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("hk_tradecal", token, **params)

def tushare_us_basic(token, **params):
    """
    美股基础信息
    
    获取美股列表信息
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("us_basic", token, **params)

def tushare_us_tradecal(token, **params):
    """
    美股交易日历
    
    获取美股交易日历信息
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("us_tradecal", token, **params)

def tushare_us_daily(token, **params):
    """
    美股日线行情
    
    获取美股行情（未复权），包括全部股票全历史行情，以及重要的市场和估值指标
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("us_daily", token, **params)

def tushare_bak_daily(token, **params):
    """
    备用行情
    
    获取备用行情，包括特定的行情指标(数据从2017年中左右开始，早期有几天数据缺失，近期正常)
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("bak_daily", token, **params)

def tushare_repo_daily(token, **params):
    """
    债券回购日行情
    
    债券回购日行情
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("repo_daily", token, **params)

def tushare_ths_index(token, **params):
    """
    同花顺行业概念板块
    
    获取同花顺板块指数。注：数据版权归属同花顺，如做商业用途，请主动联系同花顺，如需帮助请联系微信：waditu_a
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("ths_index", token, **params)

def tushare_ths_daily(token, **params):
    """
    同花顺概念和行业指数行情
    
    获取同花顺板块指数行情。注：数据版权归属同花顺，如做商业用途，请主动联系同花顺，如需帮助请联系微信：waditu_a
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("ths_daily", token, **params)

def tushare_ths_member(token, **params):
    """
    同花顺行业概念成分
    
    获取同花顺概念板块成分列表注：数据版权归属同花顺，如做商业用途，请主动联系同花顺。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("ths_member", token, **params)

def tushare_bak_basic(token, **params):
    """
    股票历史列表
    
    获取备用基础列表，数据从2016年开始
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("bak_basic", token, **params)

def tushare_fund_sales_ratio(token, **params):
    """
    各渠道公募基金销售保有规模占比
    
    获取各渠道公募基金销售保有规模占比数据，年度更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fund_sales_ratio", token, **params)

def tushare_fund_sales_vol(token, **params):
    """
    销售机构公募基金销售保有规模
    
    获取销售机构公募基金销售保有规模数据，本数据从2021年Q1开始公布，季度更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fund_sales_vol", token, **params)

def tushare_broker_recommend(token, **params):
    """
    券商月度金股
    
    获取券商月度金股，一般1日~3日内更新当月数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("broker_recommend", token, **params)

def tushare_sz_daily_info(token, **params):
    """
    深圳市场每日交易情况
    
    获取深圳市场每日交易概况
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("sz_daily_info", token, **params)

def tushare_cb_call(token, **params):
    """
    可转债赎回信息
    
    获取可转债到期赎回、强制赎回等信息。数据来源于公开披露渠道，供个人和机构研究使用，请不要用于数据商业目的。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cb_call", token, **params)

def tushare_bond_blk(token, **params):
    """
    大宗交易
    
    获取沪深交易所债券大宗交易数据，可以通过[**数据工具**](https://tushare.pro/webclient/)调试和查看数据。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("bond_blk", token, **params)

def tushare_bond_blk_detail(token, **params):
    """
    大宗交易明细
    
    获取沪深交易所债券大宗交易数据，可以通过[**数据工具**](https://tushare.pro/webclient/)调试和查看数据。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("bond_blk_detail", token, **params)

def tushare_ccass_hold_detail(token, **params):
    """
    中央结算系统持股明细
    
    获取中央结算系统机构席位持股明细，数据覆盖**全历史**，根据交易所披露时间，当日数据在下一交易日早上9点前完成
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("ccass_hold_detail", token, **params)

def tushare_stk_surv(token, **params):
    """
    机构调研数据
    
    获取上市公司机构调研记录数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_surv", token, **params)

def tushare_sge_basic(token, **params):
    """
    上海黄金基础信息
    
    获取上海黄金交易所现货合约基础信息
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("sge_basic", token, **params)

def tushare_sge_daily(token, **params):
    """
    上海黄金现货日行情
    
    获取上海黄金交易所现货合约日线行情
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("sge_daily", token, **params)

def tushare_report_rc(token, **params):
    """
    券商盈利预测数据
    
    获取券商（卖方）每天研报的盈利预测数据，数据从2010年开始，每晚19~22点更新当日数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("report_rc", token, **params)

def tushare_cyq_perf(token, **params):
    """
    每日筹码及胜率
    
    获取A股每日筹码平均成本和胜率情况，每天17~18点左右更新，数据从2018年开始
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cyq_perf", token, **params)

def tushare_cyq_chips(token, **params):
    """
    每日筹码分布
    
    获取A股每日的筹码分布情况，提供各价位占比，数据从2018年开始，每天17~18点之间更新当日数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cyq_chips", token, **params)

def tushare_ccass_hold(token, **params):
    """
    中央结算系统持股统计
    
    获取中央结算系统持股汇总数据，覆盖全部历史数据，根据交易所披露时间，当日数据在下一交易日早上9点前完成入库
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("ccass_hold", token, **params)

def tushare_stk_factor(token, **params):
    """
    股票技术面因子
    
    获取股票每日技术面因子数据，用于跟踪股票当前走势情况，数据由Tushare社区自产，覆盖全历史
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_factor", token, **params)

def tushare_limit_list_d(token, **params):
    """
    涨跌停和炸板数据
    
    获取A股每日涨跌停、炸板数据情况，数据从2020年开始（不提供ST股票的统计）
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("limit_list_d", token, **params)

def tushare_stock_mx(token, **params):
    """
    动能因子
    
    获取小佩数据动量因子数据，可以获取股票动能评级数据，包括最新及过去历史数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stock_mx", token, **params)

def tushare_stock_vx(token, **params):
    """
    估值因子
    
    小沛估值因子
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stock_vx", token, **params)

def tushare_hk_mins(token, **params):
    """
    港股分钟行情
    
    港股分钟数据，支持1min/5min/15min/30min/60min行情，提供Python SDK和 http Restful API两种方式
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("hk_mins", token, **params)

def tushare_cb_rate(token, **params):
    """
    可转债票面利率
    
    获取可转债票面利率
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cb_rate", token, **params)

def tushare_ci_daily(token, **params):
    """
    中信行业指数日行情
    
    获取中信行业指数日线行情
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("ci_daily", token, **params)

def tushare_sf_month(token, **params):
    """
    社融增量（月度）
    
    获取月度社会融资数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("sf_month", token, **params)

def tushare_hm_list(token, **params):
    """
    市场游资最全名录
    
    获取游资分类名录信息
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("hm_list", token, **params)

def tushare_hm_detail(token, **params):
    """
    游资交易每日明细
    
    获取每日游资交易明细，数据开始于2022年8。游资分类名录，请点击<a href="https://tushare.pro/document/2?doc_id=311">游资名录</a>
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("hm_detail", token, **params)

def tushare_ft_mins(token, **params):
    """
    历史分钟行情
    
    获取全市场期货合约分钟数据，支持1min/5min/15min/30min/60min行情，提供Python SDK和 http Restful API两种方式，如果需要主力合约分钟，请先通过主力[mapping](https://tushare.pro/document/2?doc_id=189)接口获取对应的合约代码后提取分钟。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("ft_mins", token, **params)

def tushare_realtime_quote(token, **params):
    """
    实时行情（爬虫）
    
    本接口是tushare org版实时接口的顺延，数据来自网络，且不进入tushare服务器，属于爬虫接口，请将tushare升级到1.3.3版本以上。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("realtime_quote", token, **params)

def tushare_realtime_tick(token, **params):
    """
    实时成交（爬虫）
    
    本接口是tushare org版实时接口的顺延，数据来自网络，且不进入tushare服务器，属于爬虫接口，数据包括该股票当日开盘以来的所有分笔成交数据。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("realtime_tick", token, **params)

def tushare_realtime_list(token, **params):
    """
    实时排名（爬虫）
    
    本接口是tushare org版实时接口的顺延，数据来自网络，且不进入tushare服务器，属于爬虫接口，数据包括该股票当日开盘以来的所有分笔成交数据。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("realtime_list", token, **params)

def tushare_ths_hot(token, **params):
    """
    同花顺App热榜数
    
    获取同花顺App热榜数据，包括热股、概念板块、ETF、可转债、港美股等等，每日盘中提取4次，收盘后4次，最晚22点提取一次。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("ths_hot", token, **params)

def tushare_dc_hot(token, **params):
    """
    东方财富App热榜
    
    获取东方财富App热榜数据，包括A股市场、ETF基金、港股市场、美股市场等等，每日盘中提取4次，收盘后4次，最晚22点提取一次。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("dc_hot", token, **params)

def tushare_bc_otcqt(token, **params):
    """
    柜台流通式债券报价
    
    柜台流通式债券报价
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("bc_otcqt", token, **params)

def tushare_bc_bestotcqt(token, **params):
    """
    柜台流通式债券最优报价
    
    柜台流通式债券最优报价
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("bc_bestotcqt", token, **params)

def tushare_cn_pmi(token, **params):
    """
    采购经理指数（PMI）
    
    采购经理人指数
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("cn_pmi", token, **params)

def tushare_margin_secs(token, **params):
    """
    融资融券标的（盘前）
    
    获取沪深京三大交易所融资融券标的（包括ETF），每天盘前更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("margin_secs", token, **params)

def tushare_sw_daily(token, **params):
    """
    申万行业指数日行情
    
    获取申万行业日线行情（默认是申万2021版行情）
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("sw_daily", token, **params)

def tushare_stk_factor_pro(token, **params):
    """
    股票技术面因子(专业版）
    
    获取股票每日技术面因子数据，用于跟踪股票当前走势情况，数据由Tushare社区自产，覆盖全历史；输出参数_bfq表示不复权，_qfq表示前复权 _hfq表示后复权，描述中说明了因子的默认传参，如需要特殊参数或者更多因子可以联系管理员评估
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_factor_pro", token, **params)

def tushare_stk_premarket(token, **params):
    """
    每日股本（盘前）
    
    每日开盘前获取当日股票的股本情况，包括总股本和流通股本，涨跌停价格等。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_premarket", token, **params)

def tushare_slb_len(token, **params):
    """
    转融资交易汇总
    
    转融通融资汇总
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("slb_len", token, **params)

def tushare_slb_sec(token, **params):
    """
    转融券交易汇总
    
    转融通转融券交易汇总
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("slb_sec", token, **params)

def tushare_slb_sec_detail(token, **params):
    """
    转融券交易明细
    
    转融券交易明细
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("slb_sec_detail", token, **params)

def tushare_slb_len_mm(token, **params):
    """
    做市借券交易汇总
    
    做市借券交易汇总
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("slb_len_mm", token, **params)

def tushare_index_member_all(token, **params):
    """
    申万行业成分（分级）
    
    按三级分类提取申万行业成分，可提供某个分类的所有成分，也可按股票代码提取所属分类，参数灵活
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("index_member_all", token, **params)

def tushare_stk_weekly_monthly(token, **params):
    """
    周/月线行情(每日更新)
    
    股票周/月线行情(每日更新)
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_weekly_monthly", token, **params)

def tushare_fut_weekly_monthly(token, **params):
    """
    期货周/月线行情(每日更新)
    
    期货周/月线行情(每日更新)
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fut_weekly_monthly", token, **params)

def tushare_us_daily_adj(token, **params):
    """
    美股复权行情
    
    获取美股复权行情，支持美股全市场股票，提供股本、市值、复权因子和成交信息等多个数据指标
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("us_daily_adj", token, **params)

def tushare_hk_daily_adj(token, **params):
    """
    港股复权行情
    
    获取港股复权行情，提供股票股本、市值和成交及换手多个数据指标
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("hk_daily_adj", token, **params)

def tushare_rt_fut_min(token, **params):
    """
    实时分钟行情
    
    获取全市场期货合约实时分钟数据，支持1min/5min/15min/30min/60min行情，提供Python SDK、 http Restful API和websocket三种方式，如果需要主力合约分钟，请先通过主力[mapping](https://tushare.pro/document/2?doc_id=189)接口获取对应的合约代码后提取分钟。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("rt_fut_min", token, **params)

def tushare_opt_mins(token, **params):
    """
    期权分钟行情
    
    获取全市场期权合约分钟数据，支持1min/5min/15min/30min/60min行情，提供Python SDK和 http Restful API两种方式。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("opt_mins", token, **params)

def tushare_moneyflow_ind_ths(token, **params):
    """
    行业资金流向（THS）
    
    获取同花顺行业板块资金流向，每日盘后更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("moneyflow_ind_ths", token, **params)

def tushare_moneyflow_ind_dc(token, **params):
    """
    板块资金流向（DC）
    
    获取东方财富板块资金流向，每天盘后更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("moneyflow_ind_dc", token, **params)

def tushare_moneyflow_mkt_dc(token, **params):
    """
    大盘资金流向（DC）
    
    获取东方财富大盘资金流向数据，每日盘后更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("moneyflow_mkt_dc", token, **params)

def tushare_kpl_list(token, **params):
    """
    榜单数据（开盘啦）
    
    获取开盘啦涨停、跌停、炸板等榜单数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("kpl_list", token, **params)

def tushare_moneyflow_ths(token, **params):
    """
    个股资金流向（THS）
    
    获取同花顺个股资金流向数据，每日盘后更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("moneyflow_ths", token, **params)

def tushare_moneyflow_dc(token, **params):
    """
    个股资金流向（DC）
    
    获取东方财富个股资金流向数据，每日盘后更新，数据开始于20230911
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("moneyflow_dc", token, **params)

def tushare_kpl_concept(token, **params):
    """
    题材数据（开盘啦）
    
    获取开盘啦概念题材列表，每天盘后更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("kpl_concept", token, **params)

def tushare_kpl_concept_cons(token, **params):
    """
    题材成分（开盘啦）
    
    获取开盘啦概念题材的成分股
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("kpl_concept_cons", token, **params)

def tushare_stk_auction_o(token, **params):
    """
    股票开盘集合竞价数据
    
    股票开盘9:30集合竞价数据，每天盘后更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_auction_o", token, **params)

def tushare_stk_auction_c(token, **params):
    """
    股票收盘集合竞价数据
    
    股票收盘15:00集合竞价数据，每天盘后更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_auction_c", token, **params)

def tushare_limit_list_ths(token, **params):
    """
    同花顺涨跌停榜单
    
    获取同花顺每日涨跌停榜单数据，历史数据从20231101开始提供，增量每天16点左右更新
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("limit_list_ths", token, **params)

def tushare_limit_step(token, **params):
    """
    涨停股票连板天梯
    
    获取每天连板个数晋级的股票，可以分析出每天连续涨停进阶个数，判断强势热度
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("limit_step", token, **params)

def tushare_limit_cpt_list(token, **params):
    """
    涨停最强板块统计
    
    获取每天涨停股票最多最强的概念板块，可以分析强势板块的轮动，判断资金动向
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("limit_cpt_list", token, **params)

def tushare_idx_factor_pro(token, **params):
    """
    指数技术面因子(专业版)
    
    获取指数每日技术面因子数据，用于跟踪指数当前走势情况，数据由Tushare社区自产，覆盖全历史；输出参数_bfq表示不复权描述中说明了因子的默认传参，如需要特殊参数或者更多因子可以联系管理员评估，指数包括大盘指数 申万行业指数 中信指数
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("idx_factor_pro", token, **params)

def tushare_fund_factor_pro(token, **params):
    """
    基金技术面因子(专业版)
    
    获取场内基金每日技术面因子数据，用于跟踪场内基金当前走势情况，数据由Tushare社区自产，覆盖全历史；输出参数_bfq表示不复权，描述中说明了因子的默认传参，如需要特殊参数或者更多因子可以联系管理员评估
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("fund_factor_pro", token, **params)

def tushare_dc_index(token, **params):
    """
    东方财富概念板块
    
    获取东方财富每个交易日的概念板块数据，支持按日期查询
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("dc_index", token, **params)

def tushare_dc_member(token, **params):
    """
    东方财富概念成分
    
    获取东方财富板块每日成分数据，可以根据概念板块代码和交易日期，获取历史成分
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("dc_member", token, **params)

def tushare_stk_nineturn(token, **params):
    """
    神奇九转指标
    
    神奇九转（又称“九转序列”）是一种基于技术分析的股票趋势反转指标，其思想来源于技术分析大师汤姆·迪马克（Tom DeMark）的TD序列。该指标的核心功能是通过识别股价在上涨或下跌过程中连续9天的特定走势，来判断股价的潜在反转点，从而帮助投资者提高抄底和逃顶的成功率，日线级别配合60min的九转效果更好，数据从20230101开始。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_nineturn", token, **params)

def tushare_stk_week_month_adj(token, **params):
    """
    周/月线复权行情(每日更新)
    
    股票周/月线行情(复权--每日更新)
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_week_month_adj", token, **params)

def tushare_ft_limit(token, **params):
    """
    期货合约涨跌停价格
    
    获取所有期货合约每天的涨跌停价格及最低保证金率，数据开始于2005年。
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("ft_limit", token, **params)

def tushare_stk_auction(token, **params):
    """
    开盘竞价成交（当日）
    
    获取当日个股和ETF的集合竞价成交情况，每天9点25后可以获取当日的集合竞价成交数据
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_auction", token, **params)

def tushare_stk_mins(token, **params):
    """
    分钟行情
    
    获取A股分钟数据，支持1min/5min/15min/30min/60min行情，提供Python SDK和 http Restful API两种方式
    
    Args:
        token (str): Tushare API令牌
        **params: API参数
    
    Returns:
        pandas.DataFrame: API返回数据
    """
    return call_tushare_api("stk_mins", token, **params)

# 函数映射字典
API_FUNCTIONS = {
    'fund_basic': tushare_fund_basic,
    'stock_basic': tushare_stock_basic,
    'trade_cal': tushare_trade_cal,
    'daily': tushare_daily,
    'adj_factor': tushare_adj_factor,
    'daily_basic': tushare_daily_basic,
    'income': tushare_income,
    'balancesheet': tushare_balancesheet,
    'cashflow': tushare_cashflow,
    'forecast': tushare_forecast,
    'express': tushare_express,
    'moneyflow_hsgt': tushare_moneyflow_hsgt,
    'hsgt_top10': tushare_hsgt_top10,
    'ggt_top10': tushare_ggt_top10,
    'margin': tushare_margin,
    'margin_detail': tushare_margin_detail,
    'top10_holders': tushare_top10_holders,
    'top10_floatholders': tushare_top10_floatholders,
    'fina_indicator': tushare_fina_indicator,
    'fina_audit': tushare_fina_audit,
    'fina_mainbz': tushare_fina_mainbz,
    'tmt_twincomedetail': tushare_tmt_twincomedetail,
    'tmt_twincome': tushare_tmt_twincome,
    'index_basic': tushare_index_basic,
    'index_daily': tushare_index_daily,
    'index_weight': tushare_index_weight,
    'namechange': tushare_namechange,
    'dividend': tushare_dividend,
    'hs_const': tushare_hs_const,
    'top_list': tushare_top_list,
    'top_inst': tushare_top_inst,
    'pledge_stat': tushare_pledge_stat,
    'pledge_detail': tushare_pledge_detail,
    'stock_company': tushare_stock_company,
    'bo_monthly': tushare_bo_monthly,
    'bo_weekly': tushare_bo_weekly,
    'bo_daily': tushare_bo_daily,
    'bo_cinema': tushare_bo_cinema,
    'fund_company': tushare_fund_company,
    'fund_nav': tushare_fund_nav,
    'fund_div': tushare_fund_div,
    'fund_portfolio': tushare_fund_portfolio,
    'new_share': tushare_new_share,
    'repurchase': tushare_repurchase,
    'concept': tushare_concept,
    'concept_detail': tushare_concept_detail,
    'fund_daily': tushare_fund_daily,
    'index_dailybasic': tushare_index_dailybasic,
    'fut_basic': tushare_fut_basic,
    'fut_daily': tushare_fut_daily,
    'fut_holding': tushare_fut_holding,
    'fut_wsr': tushare_fut_wsr,
    'fut_settle': tushare_fut_settle,
    'news': tushare_news,
    'weekly': tushare_weekly,
    'monthly': tushare_monthly,
    'shibor': tushare_shibor,
    'shibor_quote': tushare_shibor_quote,
    'shibor_lpr': tushare_shibor_lpr,
    'libor': tushare_libor,
    'hibor': tushare_hibor,
    'cctv_news': tushare_cctv_news,
    'film_record': tushare_film_record,
    'opt_basic': tushare_opt_basic,
    'opt_daily': tushare_opt_daily,
    'share_float': tushare_share_float,
    'block_trade': tushare_block_trade,
    'disclosure_date': tushare_disclosure_date,
    'stk_account': tushare_stk_account,
    'stk_account_old': tushare_stk_account_old,
    'stk_holdernumber': tushare_stk_holdernumber,
    'moneyflow': tushare_moneyflow,
    'index_weekly': tushare_index_weekly,
    'index_monthly': tushare_index_monthly,
    'wz_index': tushare_wz_index,
    'gz_index': tushare_gz_index,
    'stk_holdertrade': tushare_stk_holdertrade,
    'anns_d': tushare_anns_d,
    'fx_obasic': tushare_fx_obasic,
    'fx_daily': tushare_fx_daily,
    'teleplay_record': tushare_teleplay_record,
    'index_classify': tushare_index_classify,
    'stk_limit': tushare_stk_limit,
    'cb_basic': tushare_cb_basic,
    'cb_issue': tushare_cb_issue,
    'cb_daily': tushare_cb_daily,
    'hk_hold': tushare_hk_hold,
    'fut_mapping': tushare_fut_mapping,
    'hk_basic': tushare_hk_basic,
    'hk_daily': tushare_hk_daily,
    'stk_managers': tushare_stk_managers,
    'stk_rewards': tushare_stk_rewards,
    'major_news': tushare_major_news,
    'ggt_daily': tushare_ggt_daily,
    'ggt_monthly': tushare_ggt_monthly,
    'fund_adj': tushare_fund_adj,
    'yc_cb': tushare_yc_cb,
    'fund_share': tushare_fund_share,
    'fund_manager': tushare_fund_manager,
    'index_global': tushare_index_global,
    'suspend_d': tushare_suspend_d,
    'daily_info': tushare_daily_info,
    'fut_weekly_detail': tushare_fut_weekly_detail,
    'us_tycr': tushare_us_tycr,
    'us_trycr': tushare_us_trycr,
    'us_tbr': tushare_us_tbr,
    'us_tltr': tushare_us_tltr,
    'us_trltr': tushare_us_trltr,
    'cn_gdp': tushare_cn_gdp,
    'cn_cpi': tushare_cn_cpi,
    'eco_cal': tushare_eco_cal,
    'cn_m': tushare_cn_m,
    'cn_ppi': tushare_cn_ppi,
    'cb_price_chg': tushare_cb_price_chg,
    'cb_share': tushare_cb_share,
    'hk_tradecal': tushare_hk_tradecal,
    'us_basic': tushare_us_basic,
    'us_tradecal': tushare_us_tradecal,
    'us_daily': tushare_us_daily,
    'bak_daily': tushare_bak_daily,
    'repo_daily': tushare_repo_daily,
    'ths_index': tushare_ths_index,
    'ths_daily': tushare_ths_daily,
    'ths_member': tushare_ths_member,
    'bak_basic': tushare_bak_basic,
    'fund_sales_ratio': tushare_fund_sales_ratio,
    'fund_sales_vol': tushare_fund_sales_vol,
    'broker_recommend': tushare_broker_recommend,
    'sz_daily_info': tushare_sz_daily_info,
    'cb_call': tushare_cb_call,
    'bond_blk': tushare_bond_blk,
    'bond_blk_detail': tushare_bond_blk_detail,
    'ccass_hold_detail': tushare_ccass_hold_detail,
    'stk_surv': tushare_stk_surv,
    'sge_basic': tushare_sge_basic,
    'sge_daily': tushare_sge_daily,
    'report_rc': tushare_report_rc,
    'cyq_perf': tushare_cyq_perf,
    'cyq_chips': tushare_cyq_chips,
    'ccass_hold': tushare_ccass_hold,
    'stk_factor': tushare_stk_factor,
    'limit_list_d': tushare_limit_list_d,
    'stock_mx': tushare_stock_mx,
    'stock_vx': tushare_stock_vx,
    'hk_mins': tushare_hk_mins,
    'cb_rate': tushare_cb_rate,
    'ci_daily': tushare_ci_daily,
    'sf_month': tushare_sf_month,
    'hm_list': tushare_hm_list,
    'hm_detail': tushare_hm_detail,
    'ft_mins': tushare_ft_mins,
    'realtime_quote': tushare_realtime_quote,
    'realtime_tick': tushare_realtime_tick,
    'realtime_list': tushare_realtime_list,
    'ths_hot': tushare_ths_hot,
    'dc_hot': tushare_dc_hot,
    'bc_otcqt': tushare_bc_otcqt,
    'bc_bestotcqt': tushare_bc_bestotcqt,
    'cn_pmi': tushare_cn_pmi,
    'margin_secs': tushare_margin_secs,
    'sw_daily': tushare_sw_daily,
    'stk_factor_pro': tushare_stk_factor_pro,
    'stk_premarket': tushare_stk_premarket,
    'slb_len': tushare_slb_len,
    'slb_sec': tushare_slb_sec,
    'slb_sec_detail': tushare_slb_sec_detail,
    'slb_len_mm': tushare_slb_len_mm,
    'index_member_all': tushare_index_member_all,
    'stk_weekly_monthly': tushare_stk_weekly_monthly,
    'fut_weekly_monthly': tushare_fut_weekly_monthly,
    'us_daily_adj': tushare_us_daily_adj,
    'hk_daily_adj': tushare_hk_daily_adj,
    'rt_fut_min': tushare_rt_fut_min,
    'opt_mins': tushare_opt_mins,
    'moneyflow_ind_ths': tushare_moneyflow_ind_ths,
    'moneyflow_ind_dc': tushare_moneyflow_ind_dc,
    'moneyflow_mkt_dc': tushare_moneyflow_mkt_dc,
    'kpl_list': tushare_kpl_list,
    'moneyflow_ths': tushare_moneyflow_ths,
    'moneyflow_dc': tushare_moneyflow_dc,
    'kpl_concept': tushare_kpl_concept,
    'kpl_concept_cons': tushare_kpl_concept_cons,
    'stk_auction_o': tushare_stk_auction_o,
    'stk_auction_c': tushare_stk_auction_c,
    'limit_list_ths': tushare_limit_list_ths,
    'limit_step': tushare_limit_step,
    'limit_cpt_list': tushare_limit_cpt_list,
    'idx_factor_pro': tushare_idx_factor_pro,
    'fund_factor_pro': tushare_fund_factor_pro,
    'dc_index': tushare_dc_index,
    'dc_member': tushare_dc_member,
    'stk_nineturn': tushare_stk_nineturn,
    'stk_week_month_adj': tushare_stk_week_month_adj,
    'ft_limit': tushare_ft_limit,
    'stk_auction': tushare_stk_auction,
    'stk_mins': tushare_stk_mins,
}
