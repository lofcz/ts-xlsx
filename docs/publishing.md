# Publishing the fork

The package is `@lofcz/ts-xlsx`, published from `lofcz/ts-xlsx`.
`.github/workflows/release.yml` follows Edix's manual patch/minor/major release flow,
using this repository's pnpm toolchain and verification gates.

Run the first publish from an interactive terminal with Node 24 or newer, pnpm,
and npm 11.15 or newer:

```sh
pnpm run publish:first
```

The script opens npm's web login and leaves publication and 2FA prompts in your
terminal. It publishes the current version, then runs `npm trust github` to bind
the package to repository `lofcz/ts-xlsx`, workflow **`release.yml`**, with publish
permission. No GitHub environment is configured. It does not create an npm token
or put credentials in the repository.

The package must exist before npm accepts its trusted-publisher configuration.
If publication succeeds but trust configuration fails, rerun the script: it skips
the already-published version. An existing different trust relationship is never
revoked automatically; inspect it with `npm trust list @lofcz/ts-xlsx`.

After bootstrap, release through GitHub:

```sh
gh workflow run release.yml --repo lofcz/ts-xlsx -f release_type=patch
```

The workflow checks types, tests, build output and OOXML, bumps the version, publishes
with OIDC and provenance, pushes the release commit/tag, and creates a GitHub release.
The upstream `publish.yml` is guarded for the upstream repository and does not
publish this fork.
