-- Table: data.listing

-- DROP TABLE IF EXISTS data.listing;

CREATE TABLE IF NOT EXISTS data.listing
(
    id integer NOT NULL DEFAULT nextval('data.listing_id_seq'::regclass),
    isin character varying COLLATE pg_catalog."default",
    trading_currency character varying COLLATE pg_catalog."default",
    sedol character varying COLLATE pg_catalog."default",
    bloomberg_ticker character varying COLLATE pg_catalog."default",
    ric character varying COLLATE pg_catalog."default",
    listing_date date,
    primary_listing boolean,
    exchange character varying COLLATE pg_catalog."default",
    delisting_date date,
    shareclass_id integer,
    CONSTRAINT listing_pkey PRIMARY KEY (id)
)

TABLESPACE pg_default;

ALTER TABLE IF EXISTS data.listing
    OWNER to automations;

REVOKE ALL ON TABLE data.listing FROM factsheet;
REVOKE ALL ON TABLE data.listing FROM finance;

GRANT ALL ON TABLE data.listing TO automations;

GRANT ALL ON TABLE data.listing TO azure_pg_admin;

GRANT SELECT ON TABLE data.listing TO factsheet;

GRANT SELECT ON TABLE data.listing TO finance;