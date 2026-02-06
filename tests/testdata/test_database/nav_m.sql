-- Table: data.nav_m

-- DROP TABLE IF EXISTS data.nav_m;

CREATE TABLE IF NOT EXISTS data.nav_m
(
    id integer NOT NULL DEFAULT nextval('data.nav_m_id_seq'::regclass),
    isin character varying COLLATE pg_catalog."default",
    nav_date date,
    nav numeric,
    aum numeric,
    sharesos integer,
    ccy character varying COLLATE pg_catalog."default",
    CONSTRAINT nav_m_pkey UNIQUE (isin, nav_date, ccy)
)

TABLESPACE pg_default;

ALTER TABLE IF EXISTS data.nav_m
    OWNER to automations;

REVOKE ALL ON TABLE data.nav_m FROM factsheet;
REVOKE ALL ON TABLE data.nav_m FROM finance;
REVOKE ALL ON TABLE data.nav_m FROM marketing;

GRANT ALL ON TABLE data.nav_m TO automations;

GRANT ALL ON TABLE data.nav_m TO azure_pg_admin;

GRANT SELECT ON TABLE data.nav_m TO factsheet;

GRANT SELECT ON TABLE data.nav_m TO finance;

GRANT SELECT ON TABLE data.nav_m TO marketing;