$ids = "1jfBAVQLcc83c8XYvm5sC4EsPTTCBh3HzXarVGM3AkJcN3FJw2J2_Vw80", "1Bl5fmGE5Qqaf5QhAbxY8IlEsJKhyaOhsRWXmqN7DmYSXqUgkmHpuRVPg", "12TYbMpFRP8c2yHV-ki6IZpb46vVt95D2K64dX_Pv6P2BKxamivUOiTpI", "1y3ozDudDr7Xhk8VElvby-hmoX18p2ouKk6Bv1Y08PFgXwffPiblsYPa3", "1glCcuHbqK-gXhJvSFW2y9y9IV-rQgA7CtLa2M-YEIKEZa9b24QRl2PFr", "1nOTi--43wzeS2_WDR4V5JuCsZmKXxO7K6SxLMAKuuyNPVueg3dWePUDg", "1Oa96p_Wao_lsIGwCYuUTeexRmEj7ADLT0bAesxnrbwMoQ_9QkI26M-Yy", "1j-LNozyIk5gykPG9C-oK5hgfLskEjwf_IQxs590Q6_SnVFimHfiWePYC", "1Jshrq5uMUIMUNc3D0QIqGTB3legR84JHu1e5PAWooRorBIGfh9EgeL9G", "10zNX30ksxSW9zTUEnws3yvF9Zvg6eJ2W_aQI3h9GSlj2VQxOs9ECyUAT", "1_VhlvjGdcqGfrhRt7selU2-W4Y8zdyXOBmWJ0lioOasRPGfCVwk_zfY3", "1EHRHxkV1gGP7nWq_c9XCvHdcohEbRKRSiJwTSDNUkVIkF6mQ9PoQUYc-", "1WiZlVbHgoS__mCe_QetxKShLWYSTw1lQeUxCvFr6OwkMtzyp3qoU5v9l"

foreach ($id in $ids) {
    clasp setting scriptId $id
    clasp push
}