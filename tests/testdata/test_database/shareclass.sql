-- Table: data.shareclass

-- DROP TABLE IF EXISTS data.shareclass;

CREATE TABLE IF NOT EXISTS data.shareclass
(
    id integer NOT NULL DEFAULT nextval('data.shareclass_id_seq'::regclass),
    isin character varying COLLATE pg_catalog."default",
    share_class_name character varying COLLATE pg_catalog."default",
    gsp_account_number character varying COLLATE pg_catalog."default",
    share_class_currency character varying COLLATE pg_catalog."default",
    ter numeric,
    share_class_inception_date date,
    income_treatment character varying COLLATE pg_catalog."default",
    distribution_frequency character varying COLLATE pg_catalog."default",
    wkn character varying COLLATE pg_catalog."default",
    ise_sedol character varying COLLATE pg_catalog."default",
    valoren numeric,
    common_code numeric,
    primary_share_class boolean,
    share_class_live boolean,
    agreed_spread numeric,
    pre_us_tolerance numeric,
    post_us_tolerance numeric,
    alert_tolerance numeric,
    share_class_lifecycle character varying COLLATE pg_catalog."default",
    ref_ticker character varying COLLATE pg_catalog."default",
    srri numeric,
    fx_hedge_cost numeric,
    performance_inception_date date,
    recommended_holding_period integer,
    priip_transaction_cost numeric,
    priip_transaction_cost_date character varying COLLATE pg_catalog."default",
    entry_fee numeric,
    exit_fee numeric,
    performance_fee numeric,
    display_currency character varying COLLATE pg_catalog."default",
    product_id integer,
    CONSTRAINT shareclass_pkey PRIMARY KEY (id)
)

TABLESPACE pg_default;

ALTER TABLE IF EXISTS data.shareclass
    OWNER to automations;

REVOKE ALL ON TABLE data.shareclass FROM factsheet;
REVOKE ALL ON TABLE data.shareclass FROM finance;

GRANT ALL ON TABLE data.shareclass TO automations;

GRANT ALL ON TABLE data.shareclass TO azure_pg_admin;

GRANT SELECT ON TABLE data.shareclass TO factsheet;

GRANT SELECT ON TABLE data.shareclass TO finance;