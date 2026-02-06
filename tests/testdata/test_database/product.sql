-- Table: data.product

-- DROP TABLE IF EXISTS data.product;

CREATE TABLE IF NOT EXISTS data.product
(
    id integer NOT NULL DEFAULT nextval('data.product_id_seq'::regclass),
    client character varying COLLATE pg_catalog."default",
    sub_fund_name character varying COLLATE pg_catalog."default",
    base_currency character varying COLLATE pg_catalog."default",
    jpm_acct_no character varying COLLATE pg_catalog."default",
    sub_fund_lei character varying COLLATE pg_catalog."default",
    issuer_lei character varying COLLATE pg_catalog."default",
    approval_date date,
    asset_class character varying COLLATE pg_catalog."default",
    tracking_method character varying COLLATE pg_catalog."default",
    fund_reporting_status boolean,
    isa_eligible boolean,
    year_end character varying COLLATE pg_catalog."default",
    active_fund boolean,
    settlement_for_creation character varying COLLATE pg_catalog."default",
    settlement_for_redemption character varying COLLATE pg_catalog."default",
    valuation_point character varying COLLATE pg_catalog."default",
    nav_pricing_methodology character varying COLLATE pg_catalog."default",
    sfdr character varying COLLATE pg_catalog."default",
    inav_ric character varying COLLATE pg_catalog."default",
    inav_bbg_ticker character varying COLLATE pg_catalog."default",
    inav_wkn character varying COLLATE pg_catalog."default",
    inav_isin character varying COLLATE pg_catalog."default",
    inav_currency character varying COLLATE pg_catalog."default",
    inav_start_date character varying COLLATE pg_catalog."default",
    cu_creation_fee character varying COLLATE pg_catalog."default",
    cu_redemption_fee character varying COLLATE pg_catalog."default",
    cu_size_units numeric,
    ap_dealing_cutoff_time character varying COLLATE pg_catalog."default",
    ap_dealing_cutoff_day character varying COLLATE pg_catalog."default",
    createredeem_method character varying COLLATE pg_catalog."default",
    index_net_total_return_ticker character varying COLLATE pg_catalog."default",
    im character varying COLLATE pg_catalog."default",
    sub_advisor character varying COLLATE pg_catalog."default",
    im_start_date character varying COLLATE pg_catalog."default",
    lead_market_maker character varying COLLATE pg_catalog."default",
    issuer character varying COLLATE pg_catalog."default",
    manco character varying COLLATE pg_catalog."default",
    gsp_account_number character varying COLLATE pg_catalog."default",
    domicile character varying COLLATE pg_catalog."default",
    trustee character varying COLLATE pg_catalog."default",
    depositary character varying COLLATE pg_catalog."default",
    administrator character varying COLLATE pg_catalog."default",
    registrar character varying COLLATE pg_catalog."default",
    custodian character varying COLLATE pg_catalog."default",
    ucits_compliant boolean,
    ucits_eligible boolean,
    sipp_eligible boolean,
    common_code numeric,
    live boolean,
    synthetic boolean,
    swap_fee numeric,
    swap_counterparty character varying COLLATE pg_catalog."default",
    deregistration_date numeric,
    liquidation_date date,
    fsin character varying COLLATE pg_catalog."default",
    cfi character varying COLLATE pg_catalog."default",
    id_amf numeric,
    product_class character varying COLLATE pg_catalog."default",
    website_product_number numeric,
    jv_percent numeric,
    shariah_compliant boolean,
    bbg_etc_underlying character varying COLLATE pg_catalog."default",
    fund_cbi_code character varying COLLATE pg_catalog."default",
    monthly_performance_grouping character varying COLLATE pg_catalog."default",
    common_depositary character varying COLLATE pg_catalog."default",
    storage_location character varying COLLATE pg_catalog."default",
    storage_inspector character varying COLLATE pg_catalog."default",
    storage_type character varying COLLATE pg_catalog."default",
    pea_eligible boolean,
    entitlement_asset character varying COLLATE pg_catalog."default",
    entitlement_class character varying COLLATE pg_catalog."default",
    entitlement_unit character varying COLLATE pg_catalog."default",
    securities_lending boolean,
    etc_strategy character varying COLLATE pg_catalog."default",
    etc_staking_reward_net character varying COLLATE pg_catalog."default",
    etc_staking_service_fee numeric,
    etc_staking_provider character varying COLLATE pg_catalog."default",
    paying_agent character varying COLLATE pg_catalog."default",
    fiscal_agent character varying COLLATE pg_catalog."default",
    fx_hedge_counterparty character varying COLLATE pg_catalog."default",
    sec_lending_limit numeric,
    sharpe_ratio numeric,
    sortino_ratio numeric,
    effective_duration numeric,
    mountain_view_holdingname character varying COLLATE pg_catalog."default",
    mountain_view_security_id character varying COLLATE pg_catalog."default",
    fenion_fund_sub_category character varying COLLATE pg_catalog."default",
    has_derivatives numeric,
    emt_benchmark character varying COLLATE pg_catalog."default",
    CONSTRAINT product_pkey PRIMARY KEY (id)
)

TABLESPACE pg_default;

ALTER TABLE IF EXISTS data.product
    OWNER to automations;

REVOKE ALL ON TABLE data.product FROM factsheet;
REVOKE ALL ON TABLE data.product FROM finance;

GRANT ALL ON TABLE data.product TO automations;

GRANT ALL ON TABLE data.product TO azure_pg_admin;

GRANT SELECT ON TABLE data.product TO factsheet;

GRANT SELECT ON TABLE data.product TO finance;