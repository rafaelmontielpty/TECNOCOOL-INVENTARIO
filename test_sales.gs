function testSale() {
  registerSale({
    sale_id: "SALE-" + new Date().getTime(),
    sale_type: "Mostrador",
    customer_id: "CUST-001",
    product_id: "SKU-001",
    quantity: 2,
    warehouse_id: "ALM-01",
    user: "Rafael",
    created_at: new Date()
  });
}
