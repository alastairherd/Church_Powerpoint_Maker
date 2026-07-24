use anyhow::{bail, Context, Result};
use argon2::password_hash::SaltString;
use argon2::{Argon2, PasswordHasher};
use std::io::{Read, Write};

fn main() -> Result<()> {
    eprint!("Password: ");
    std::io::stderr().flush()?;

    let mut password = String::new();
    std::io::stdin()
        .read_to_string(&mut password)
        .context("failed to read the password from standard input")?;
    let password = password.trim_end_matches(['\r', '\n']);
    if password.is_empty() {
        bail!("password cannot be empty");
    }

    let mut salt_bytes = [0_u8; 16];
    std::fs::File::open("/dev/urandom")
        .context("failed to open the operating system random source")?
        .read_exact(&mut salt_bytes)
        .context("failed to generate a password salt")?;
    let salt = SaltString::encode_b64(&salt_bytes)
        .map_err(|error| anyhow::anyhow!("failed to encode password salt: {error}"))?;
    let hash = Argon2::default()
        .hash_password(password.as_bytes(), &salt)
        .map_err(|error| anyhow::anyhow!("failed to hash password: {error}"))?;

    println!("{hash}");
    Ok(())
}
