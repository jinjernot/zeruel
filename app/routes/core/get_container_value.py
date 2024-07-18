def get_container_value(df, sku, container_name):
    """Look for the container value from df2"""

    container = df[(df["SKU"] == sku) & (df["ContainerName"] == container_name)]
    values = container["ContainerValue"].values.tolist()
    return values[0] if values else None