# Argentum Online
# https://ao-libre.github.io/ao-website/
# Fixture for the database creation
# Created on September 17th 2018
# By Juan Andres Dalmasso (CHOTS)
# Last modification: 17/09/2018 (CHOTS)

CREATE TABLE account (
    id MEDIUMINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(50) NOT NULL,
    password VARCHAR(50) NOT NULL,
    salt VARCHAR(50) NOT NULL,
    date_created TIMESTAMP NOT NULL,
    date_last_login TIMESTAMP NOT NULL,
    last_ip VARCHAR(16)
);

CREATE TABLE user (
    id MEDIUMINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    account_id MEDIUMINT UNSIGNED NOT NULL,
    
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
    bank_gold INT UNSIGNED NOT NULL,
    free_skillpoints SMALLINT UNSIGNED NOT NULL,
    assigned_skillpoints SMALLINT UNSIGNED NOT NULL,
    pet_amount TINYINT UNSIGNED NOT NULL,

    # POS
    pos_map TINYINT UNSIGNED NOT NULL,
    pos_x TINYINT UNSIGNED NOT NULL,
    pos_y TINYINT UNSIGNED NOT NULL,
    last_map TINYINT UNSIGNED NOT NULL,

    # INVENTORY
    body_id SMALLINT UNSIGNED NOT NULL,
    head_id SMALLINT UNSIGNED NOT NULL,
    weapon_id SMALLINT UNSIGNED NOT NULL,
    helmet_id SMALLINT UNSIGNED NOT NULL,
    shield_id SMALLINT UNSIGNED NOT NULL,
    items_amount TINYINT UNSIGNED NOT NULL,
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
    killed_npcs SMALLINT UNSIGNED NOT NULL,
    killed_users SMALLINT UNSIGNED NOT NULL,

    # REPUTACION
    rep_asesino SMALLINT UNSIGNED NOT NULL,
    rep_bandido SMALLINT UNSIGNED NOT NULL,
    rep_burgues SMALLINT UNSIGNED NOT NULL,
    rep_ladron SMALLINT UNSIGNED NOT NULL,
    rep_noble SMALLINT UNSIGNED NOT NULL,
    rep_plebe SMALLINT UNSIGNED NOT NULL,

    # FLAGS
    is_connected BOOLEAN NOT NULL,
    is_naked BOOLEAN NOT NULL,
    is_poisoned BOOLEAN NOT NULL,
    is_hidden BOOLEAN NOT NULL,
    is_hungry BOOLEAN NOT NULL,
    is_thirsty BOOLEAN NOT NULL,
    is_ban BOOLEAN NOT NULL,
    is_dead BOOLEAN NOT NULL,
    is_sailing BOOLEAN NOT NULL,

    # COUNTERS
    counter_pena TINYINT UNSIGNED NOT NULL,
    counter_connected INT UNSIGNED NOT NULL,
    counter_training INT UNSIGNED NOT NULL,

    # FACCION
    pertenece_consejo_real BOOLEAN NOT NULL,
    pertenece_consejo_caos BOOLEAN NOT NULL,
    pertenece_real BOOLEAN NOT NULL,
    pertenece_caos BOOLEAN NOT NULL,
    ciudadanos_matados SMALLINT UNSIGNED NOT NULL,
    criminales_matados SMALLINT UNSIGNED NOT NULL,
    recibio_armadura_real BOOLEAN NOT NULL,
    recibio_armadura_caos BOOLEAN NOT NULL,
    recibio_exp_real BOOLEAN NOT NULL,
    recibio_exp_caos BOOLEAN NOT NULL,
    recompensas_real TINYINT UNSIGNED,
    recompensas_caos TINYINT UNSIGNED,
    reenlistadas SMALLINT UNSIGNED,
    fecha_ingreso TIMESTAMP,
    nivel_ingreso SMALLINT UNSIGNED NOT NULL,
    matados_ingreso SMALLINT UNSIGNED NOT NULL,
    siguiente_recompensa SMALLINT UNSIGNED NOT NULL,

    # GUILD
    guild_index SMALLINT UNSIGNED,
    guild_aspirant_index SMALLINT UNSIGNED,
    guild_member_of VARCHAR(50),
    guild_member_history VARCHAR(1024),
    guild_rejected_because VARCHAR(255),

    CONSTRAINT fk_user_account FOREIGN KEY (account_id) REFERENCES account(id)
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
    amount SMALLINT UNSIGNED,
    exp INT UNSIGNED UNSIGNED NOT NULL,
    elu INT UNSIGNED UNSIGNED NOT NULL,

    PRIMARY KEY (user_id, number),
    CONSTRAINT fk_skillpoint_user FOREIGN KEY (user_id) REFERENCES user(id)
);

# Credentials for testing environment
# argentumonlinefree.cjq47ruczip9.sa-east-1.rds.amazonaws.com
# argentumonlinefree
# z7vW5jWuMkytBzuSBteRKXnXELUcEgt9
# Linux command:
# mysql -h 'argentumonlinefree.cjq47ruczip9.sa-east-1.rds.amazonaws.com' -u 'argentumonline' -p