CREATE DATABASE star_interactive;

USE star_interactive;

CREATE TABLE lean_supply
(id INT NOT NULL AUTO_INCREMENT,
sku VARCHAR(255),
item_desc VARCHAR(255),
upc BIGINT,
PRIMARY KEY (id));
