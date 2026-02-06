-- Table: data.nav

-- DROP TABLE IF EXISTS data.nav;

CREATE TABLE IF NOT EXISTS data.nav
(
    id integer NOT NULL DEFAULT nextval('data.nav_id_seq'::regclass),
    isin character varying COLLATE pg_catalog."default",
    nav_date date,
    nav numeric,
    aum numeric,
    sharesos integer,
    CONSTRAINT nav_pkey UNIQUE (isin, nav_date)
)

TABLESPACE pg_default;

ALTER TABLE IF EXISTS data.nav
    OWNER to automations;

REVOKE ALL ON TABLE data.nav FROM factsheet;
REVOKE ALL ON TABLE data.nav FROM finance;
REVOKE ALL ON TABLE data.nav FROM marketing;

GRANT ALL ON TABLE data.nav TO automations;

GRANT ALL ON TABLE data.nav TO azure_pg_admin;

GRANT SELECT ON TABLE data.nav TO factsheet;

GRANT SELECT ON TABLE data.nav TO finance;

GRANT SELECT ON TABLE data.nav TO marketing;