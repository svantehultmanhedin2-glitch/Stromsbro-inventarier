import Pill from "../common/Pill";
import PrimaryButton from "../common/PrimaryButton";


export default function LagerView({
  produkter,
  filtreradeProdukter,
  status,
  getQty,
  setQty,
  canUtlamna,
  canEditHuvudlager,
  selectMode,
  selectedIds,
  toggleSelected,
  openTx,
  openBatch,
  selectAllVisible,
  clearSelected,
openAddProd,

}) {
  return (
                <>
                  {selectMode && canUtlamna ? (
                    <div className="massBar">
                      <div className="massBar__left">
                        <div className="massBar__title">Massläge</div>
                        <div className="massBar__sub">{selectedCount} markerade</div>
                      </div>
                      <div className="massBar__actions">
                        <PrimaryButton tone="ghost" onClick={selectAllVisible}>
                          Markera alla
                        </PrimaryButton>
                        <PrimaryButton tone="ghost" onClick={clearSelected} disabled={selectedCount === 0}>
                          Rensa
                        </PrimaryButton>
                        <PrimaryButton tone="danger" onClick={() => openBatch("uttag")} disabled={selectedCount === 0}>
                          ➖ Uttag
                        </PrimaryButton>
                        <PrimaryButton tone="ok" onClick={() => openBatch("in")} disabled={selectedCount === 0}>
                          ➕ In
                        </PrimaryButton>
                      </div>
                    </div>
                  ) : null}

                  {/* ✅ NY: Lägg till produkt under Lager */}
                  {canEditHuvudlager && !selectMode ? (
                    <section className="card" style={{ marginBottom: 12 }}>
                      <div className="card__top">
                        <div className="card__title">➕ Lägg till produkt</div>
                        <Pill tone="neutral">Huvudlager</Pill>
                      </div>
                      <div className="muted" style={{ marginTop: 8 }}>
                        Skapa ny artikel i huvudlagret (påverkar inköp automatiskt).
                      </div>
                      <div className="btnRow" style={{ marginTop: 10, gridTemplateColumns: "1fr" }}>
                        <PrimaryButton tone="primary" onClick={openAddProd}>
                          ➕ Ny produkt
                        </PrimaryButton>
                      </div>
                    </section>
                  ) : null}

                  <div className="grid">
                    {filtreradeProdukter.length === 0 ? <div className="empty">Inga produkter matchar sökningen.</div> : null}

                    {filtreradeProdukter.map((p) => {
                      const st = status(p);
                      const qty = getQty(p.id);
                      const selected = selectMode && selectedIds.has(p.id);

                      return (
                        <section
                          className={`card ${selected ? "card--selected" : ""}`}
                          key={p.id}
                          onClick={() => {
                            if (!selectMode) return;
                            toggleSelected(p.id);
                          }}
                        >
                          <div className="card__top">
                            <div className="card__title">{p.produkt}</div>
                            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                              {selectMode && canUtlamna ? (
                                <label className="selectCtl" title="Markera" onClick={(e) => e.stopPropagation()}>
                                  <input type="checkbox" checked={selectedIds.has(p.id)} onChange={() => toggleSelected(p.id)} />
                                  <span>Markera</span>
                                </label>
                              ) : null}
                              <Pill tone={st.tone}>{st.text}</Pill>
                            </div>
                          </div>

                          <div className="meta">
                            <div className="meta__row">
                              <span className="meta__label">Grupp</span>
                              <span className="meta__value">{p.huvudgrupp || "—"}</span>
                            </div>
                            <div className="meta__row">
                              <span className="meta__label">Plats</span>
                              <span className="meta__value">{p.lagerplats || "—"}</span>
                            </div>
                          </div>

                          <div className="qtyRow">
                            <div className="qty">
                              <div className="qty__label">I lager</div>
                              <div className="qty__value">{toInt(p.antal, 0)} st</div>
                            </div>
                            <div className="miniMeta">
                              <div>
                                Min: <strong>{toInt(p.minAntal, 0)}</strong>
                              </div>
                              <div>
                                Beställ: <strong>{toInt(p.beställningspunkt, 0)}</strong>
                              </div>
                            </div>
                          </div>

                          <div className="qtyPick" onClick={(e) => e.stopPropagation()}>
                            <div className="qtyPick__label">Antal (st)</div>
                            <input
                              className="qtyInput"
                              type="number"
                              min={1}
                              inputMode="numeric"
                              value={qty}
                              onChange={(e) => setQty(p.id, e.target.value)}
                            />
                            <div className="qtyChips">
                              <button className="chip" type="button" onClick={() => setQty(p.id, 1)}>
                                1
                              </button>
                              <button className="chip" type="button" onClick={() => setQty(p.id, 5)}>
                                5
                              </button>
                              <button className="chip" type="button" onClick={() => setQty(p.id, 10)}>
                                10
                              </button>
                            </div>
                          </div>

                          <div className="btnRow" onClick={(e) => e.stopPropagation()}>
                            <PrimaryButton onClick={() => openTx("uttag", p)} tone="danger" disabled={!canUtlamna || selectMode}>
                              ➖ Uttag
                            </PrimaryButton>
                            <PrimaryButton onClick={() => openTx("in", p)} tone="ok" disabled={!canEditHuvudlager || selectMode}>
                              ➕ In
                            </PrimaryButton>
                          </div>

                          {isLedare ? <div className="muted" style={{ marginTop: 8 }}>Ledare kan inte göra uttag/inleverans i huvudlager.</div> : null}
                        </section>
                      );
                    })}
                  </div>
                </>
              );
}