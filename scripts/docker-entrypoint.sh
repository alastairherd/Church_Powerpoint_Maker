#!/bin/sh
set -eu

# Docker can start before the host's Tailscale resolver is available. Keep its
# resolver for internal names, but add public fallbacks if R2 cannot resolve.
# Set APP_DNS_FALLBACKS="" to disable, or supply space-separated resolver IPs.
if [ "${OBJECT_STORE:-r2}" != memory ] && [ -n "${R2_ACCOUNT_ID:-}" ]; then
    r2_host="${R2_ACCOUNT_ID}.r2.cloudflarestorage.com"
    if ! timeout 10 getent ahosts "$r2_host" >/dev/null 2>&1; then
        fallbacks=${APP_DNS_FALLBACKS-'1.1.1.1 9.9.9.9'}
        if [ -n "$fallbacks" ]; then
            for resolver in $fallbacks; do
                case "$resolver" in
                    *[!0-9a-fA-F.:]*|'') echo 'Invalid APP_DNS_FALLBACKS resolver IP' >&2; exit 1 ;;
                esac
            done
            backup=$(mktemp)
            cp /etc/resolv.conf "$backup"
            # Put fallbacks first: libc uses at most three nameservers.
            {
                for resolver in $fallbacks; do printf 'nameserver %s\n' "$resolver"; done
                cat "$backup"
            } > /etc/resolv.conf
            if timeout 10 getent ahosts "$r2_host" >/dev/null 2>&1; then
                echo 'R2 DNS recovered using fallback resolvers.' >&2
            else
                cat "$backup" > /etc/resolv.conf
                echo 'R2 DNS lookup failed, including fallback resolvers; check container networking.' >&2
            fi
            rm -f "$backup"
        else
            echo 'R2 DNS lookup failed; DNS fallbacks are disabled.' >&2
        fi
    fi
fi

exec "$@"
