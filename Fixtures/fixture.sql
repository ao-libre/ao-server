# Argentum Online
# https://ao-libre.github.io/ao-website/
# Fixture for the database creation
# Created on September 17th 2018
# By Juan Andres Dalmasso (CHOTS)
# Last modification: 10/10/2018 (CHOTS)

CREATE TABLE account (
    id MEDIUMINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(50) NOT NULL,
    password VARCHAR(64) NOT NULL,
    salt VARCHAR(10) NOT NULL,
    hash VARCHAR(32) NOT NULL,
    date_created TIMESTAMP NOT NULL,
    date_last_login TIMESTAMP NOT NULL,
    last_ip VARCHAR(16)
);

CREATE TABLE user (
    id MEDIUMINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    account_id MEDIUMINT UNSIGNED NOT NULL,
    deleted BOOLEAN NOT NULL DEFAULT FALSE,

    # INIT
    name VARCHAR(30) NOT NULL,
    level SMALLINT UNSIGNED NOT NULL,
    exp INT UNSIGNED NOT NULL,
    elu INT UNSIGNED NOT NULL,
    genre_id TINYINT UNSIGNED NOT NULL,
    race_id TINYINT UNSIGNED NOT NULL,
    class_id TINYINT UNSIGNED NOT NULL,
    home_id TINYINT UNSIGNED NOT NULL,
    description VARCHAR(255),
    gold INT UNSIGNED NOT NULL,
    bank_gold INT UNSIGNED NOT NULL DEFAULT 0,
    free_skillpoints SMALLINT UNSIGNED NOT NULL,
    assigned_skillpoints SMALLINT UNSIGNED NOT NULL,
    pet_amount TINYINT UNSIGNED NOT NULL DEFAULT 0,
    votes_amount SMALLINT UNSIGNED DEFAULT 0,

    # POS
    pos_map SMALLINT UNSIGNED NOT NULL,
    pos_x TINYINT UNSIGNED NOT NULL,
    pos_y TINYINT UNSIGNED NOT NULL,
    last_map TINYINT UNSIGNED NOT NULL DEFAULT 1,

    # INVENTORY
    body_id SMALLINT UNSIGNED NOT NULL,
    head_id SMALLINT UNSIGNED NOT NULL,
    weapon_id SMALLINT UNSIGNED NOT NULL,
    helmet_id SMALLINT UNSIGNED NOT NULL,
    shield_id SMALLINT UNSIGNED NOT NULL,
    heading TINYINT UNSIGNED NOT NULL DEFAULT 3,
    items_amount TINYINT UNSIGNED NOT NULL,
    slot_armour TINYINT UNSIGNED,
    slot_weapon TINYINT UNSIGNED,
    slot_helmet TINYINT UNSIGNED,
    slot_shield TINYINT UNSIGNED,
    slot_ammo TINYINT UNSIGNED,
    slot_ship TINYINT UNSIGNED,
    slot_ring TINYINT UNSIGNED,
    slot_bag TINYINT UNSIGNED,

    # STATS
    min_hp SMALLINT UNSIGNED NOT NULL,
    max_hp SMALLINT UNSIGNED NOT NULL,
    min_man SMALLINT UNSIGNED NOT NULL,
    max_man SMALLINT UNSIGNED NOT NULL,
    min_sta SMALLINT UNSIGNED NOT NULL,
    max_sta SMALLINT UNSIGNED NOT NULL,
    min_ham SMALLINT UNSIGNED NOT NULL,
    max_ham SMALLINT UNSIGNED NOT NULL,
    min_sed SMALLINT UNSIGNED NOT NULL,
    max_sed SMALLINT UNSIGNED NOT NULL,
    min_hit SMALLINT UNSIGNED NOT NULL,
    max_hit SMALLINT UNSIGNED NOT NULL,
    killed_npcs SMALLINT UNSIGNED NOT NULL DEFAULT 0,
    killed_users SMALLINT UNSIGNED NOT NULL DEFAULT 0,

    # REPUTACION
    rep_asesino MEDIUMINT UNSIGNED NOT NULL DEFAULT 0,
    rep_bandido MEDIUMINT UNSIGNED NOT NULL DEFAULT 0,
    rep_burgues MEDIUMINT UNSIGNED NOT NULL DEFAULT 0,
    rep_ladron MEDIUMINT UNSIGNED NOT NULL DEFAULT 0,
    rep_noble MEDIUMINT UNSIGNED NOT NULL,
    rep_plebe MEDIUMINT UNSIGNED NOT NULL,
    rep_average MEDIUMINT NOT NULL,

    # FLAGS
    is_naked BOOLEAN NOT NULL DEFAULT FALSE,
    is_poisoned BOOLEAN NOT NULL DEFAULT FALSE,
    is_hidden BOOLEAN NOT NULL DEFAULT FALSE,
    is_hungry BOOLEAN NOT NULL DEFAULT FALSE,
    is_thirsty BOOLEAN NOT NULL DEFAULT FALSE,
    is_ban BOOLEAN NOT NULL DEFAULT FALSE,
    is_dead BOOLEAN NOT NULL DEFAULT FALSE,
    is_sailing BOOLEAN NOT NULL DEFAULT FALSE,
    is_paralyzed BOOLEAN NOT NULL DEFAULT FALSE,
    is_logged BOOLEAN NOT NULL DEFAULT FALSE,

    # COUNTERS
    counter_pena SMALLINT UNSIGNED NOT NULL DEFAULT 0,
    counter_connected INT UNSIGNED NOT NULL DEFAULT 0,
    counter_training INT UNSIGNED NOT NULL DEFAULT 0,

    # FACCION
    pertenece_consejo_real BOOLEAN NOT NULL DEFAULT FALSE,
    pertenece_consejo_caos BOOLEAN NOT NULL DEFAULT FALSE,
    pertenece_real BOOLEAN NOT NULL DEFAULT FALSE,
    pertenece_caos BOOLEAN NOT NULL DEFAULT FALSE,
    ciudadanos_matados SMALLINT UNSIGNED NOT NULL DEFAULT 0,
    criminales_matados SMALLINT UNSIGNED NOT NULL DEFAULT 0,
    recibio_armadura_real BOOLEAN NOT NULL DEFAULT FALSE,
    recibio_armadura_caos BOOLEAN NOT NULL DEFAULT FALSE,
    recibio_exp_real BOOLEAN NOT NULL DEFAULT FALSE,
    recibio_exp_caos BOOLEAN NOT NULL DEFAULT FALSE,
    recompensas_real TINYINT UNSIGNED DEFAULT 0,
    recompensas_caos TINYINT UNSIGNED DEFAULT 0,
    reenlistadas SMALLINT UNSIGNED NOT NULL DEFAULT 0,
    fecha_ingreso TIMESTAMP,
    nivel_ingreso SMALLINT UNSIGNED,
    matados_ingreso SMALLINT UNSIGNED,
    siguiente_recompensa SMALLINT UNSIGNED,

    # GUILD
    guild_index SMALLINT UNSIGNED DEFAULT 0,
    guild_aspirant_index SMALLINT UNSIGNED,
    guild_member_history VARCHAR(1024),
    guild_requests_history VARCHAR(1024),
    guild_rejected_because VARCHAR(255),

    CONSTRAINT fk_user_account FOREIGN KEY (account_id) REFERENCES account(id),
    INDEX (name)
);

CREATE TABLE spell (
    user_id MEDIUMINT UNSIGNED NOT NULL,
    number TINYINT UNSIGNED NOT NULL,
    spell_id SMALLINT UNSIGNED,

    PRIMARY KEY (user_id, number),
    CONSTRAINT fk_spell_user FOREIGN KEY (user_id) REFERENCES user(id)
);

CREATE TABLE pet (
    user_id MEDIUMINT UNSIGNED NOT NULL,
    number TINYINT UNSIGNED NOT NULL,
    pet_id SMALLINT UNSIGNED,

    PRIMARY KEY (user_id, number),
    CONSTRAINT fk_pet_user FOREIGN KEY (user_id) REFERENCES user(id)
);

CREATE TABLE attribute (
    user_id MEDIUMINT UNSIGNED NOT NULL,
    number TINYINT UNSIGNED NOT NULL,
    value TINYINT UNSIGNED NOT NULL,

    PRIMARY KEY (user_id, number),
    CONSTRAINT fk_attribute_user FOREIGN KEY (user_id) REFERENCES user(id)
);

CREATE TABLE punishment (
    user_id MEDIUMINT UNSIGNED NOT NULL,
    number TINYINT UNSIGNED NOT NULL,
    reason VARCHAR(255) NOT NULL,

    PRIMARY KEY (user_id, number),
    CONSTRAINT fk_punishment_user FOREIGN KEY (user_id) REFERENCES user(id)
);

CREATE TABLE inventory_item (
    user_id MEDIUMINT UNSIGNED NOT NULL,
    number TINYINT UNSIGNED NOT NULL,
    item_id SMALLINT UNSIGNED,
    amount SMALLINT UNSIGNED,
    is_equipped BOOLEAN,

    PRIMARY KEY (user_id, number),
    CONSTRAINT fk_inventory_user FOREIGN KEY (user_id) REFERENCES user(id)
);

CREATE TABLE bank_item (
    user_id MEDIUMINT UNSIGNED NOT NULL,
    number TINYINT UNSIGNED NOT NULL,
    item_id SMALLINT UNSIGNED,
    amount SMALLINT UNSIGNED,

    PRIMARY KEY (user_id, number),
    CONSTRAINT fk_bank_user FOREIGN KEY (user_id) REFERENCES user(id)
);

CREATE TABLE skillpoint (
    user_id MEDIUMINT UNSIGNED NOT NULL,
    number TINYINT UNSIGNED NOT NULL,
    value TINYINT UNSIGNED NOT NULL,
    exp INT UNSIGNED NOT NULL,
    elu INT UNSIGNED NOT NULL,

    PRIMARY KEY (user_id, number),
    CONSTRAINT fk_skillpoint_user FOREIGN KEY (user_id) REFERENCES user(id)
);

# Credentials for testing environment
# argentumonlinefree.cjq47ruczip9.sa-east-1.rds.amazonaws.com
# argentumonline
# z7vW5jWuMkytBzuSBteRKXnXELUcEgt9
# Linux command:
# mysql -h 'argentumonlinefree.cjq47ruczip9.sa-east-1.rds.amazonaws.com' -u 'argentumonline' -p
# https://dev.mysql.com/downloads/connector/odbc/5.1.html