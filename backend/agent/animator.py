from pptx.oxml.ns import qn
from lxml import etree

def add_fade_animation(shape):
    """Adds a fade-in entrance animation to a shape."""
    sp_tree = shape._element.getparent()
    timing = sp_tree.getparent().getparent()

    # Build the animation XML
    anim_xml = f"""
    <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:tnLst>
        <p:par>
          <p:cTn id="1" dur="indefinite" restart="whenNotActive" nodeType="tmRoot">
            <p:childTnLst>
              <p:par>
                <p:cTn id="2" fill="hold">
                  <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
                  <p:childTnLst>
                    <p:par>
                      <p:cTn id="3" presetID="10" presetClass="entr" presetSubtype="0"
                             fill="hold" grpId="0" nodeType="clickEffect">
                        <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                        <p:childTnLst>
                          <p:set>
                            <p:cBhvr><p:cTn id="4" dur="1" fill="hold"/>
                              <p:tgtEl><p:spTgt spid="{shape.shape_id}"/></p:tgtEl>
                              <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                            </p:cBhvr>
                            <p:to><p:strVal val="visible"/></p:to>
                          </p:set>
                          <p:animEffect transition="in" filter="fade">
                            <p:cBhvr><p:cTn id="5" dur="500"/>
                              <p:tgtEl><p:spTgt spid="{shape.shape_id}"/></p:tgtEl>
                            </p:cBhvr>
                          </p:animEffect>
                        </p:childTnLst>
                      </p:cTn>
                    </p:par>
                  </p:childTnLst>
                </p:cTn>
              </p:par>
            </p:childTnLst>
          </p:cTn>
        </p:par>
      </p:tnLst>
    </p:timing>
    """
    try:
        timing_elem = etree.fromstring(anim_xml)
        slide_elem = shape._element.getparent().getparent()
        existing = slide_elem.find(qn('p:timing'))
        if existing is not None:
            slide_elem.remove(existing)
        slide_elem.append(timing_elem)
    except Exception:
        pass  # Silently skip if animation injection fails