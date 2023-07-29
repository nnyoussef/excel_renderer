package nnyo.excel.renderer;

import com.steadystate.css.parser.CSSOMParser;
import com.steadystate.css.parser.SACParserCSS3;
import nnyo.excel.renderer.constantes.CssConstantes;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.css.sac.InputSource;
import org.w3c.dom.css.CSSRule;
import org.w3c.dom.css.CSSRuleList;
import org.w3c.dom.css.CSSStyleDeclaration;
import org.w3c.dom.css.CSSStyleSheet;

import java.awt.*;
import java.io.IOException;
import java.io.StringReader;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import static java.lang.Integer.parseInt;
import static java.util.Optional.ofNullable;
import static java.util.stream.Collectors.joining;
import static nnyo.excel.renderer.constantes.CssConstantes.*;
import static org.apache.commons.lang3.StringUtils.EMPTY;
import static org.apache.commons.lang3.StringUtils.SPACE;
import static org.apache.poi.ss.usermodel.BorderStyle.NONE;
import static org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT;
import static org.apache.poi.ss.usermodel.VerticalAlignment.CENTER;

public class CellStyleProcessor {

    private final Map<String, XSSFCellStyle> cacheCellStyle = new HashMap<>(30);

    private XSSFWorkbook xssfWorkbook;

    private Map<String, Map<String, String>> cssRuleDeclaration = new HashMap<>(30);

    //-------------------------- Init One Time - Public Usage -------------------------------------------------
    public static CellStyleProcessor init(String css,
                                          XSSFWorkbook xssfWorkbook) throws IOException {

        Map<String, Map<String, String>> cssRuleDeclaration = new HashMap<>(30);

        InputSource inputSource = new InputSource(new StringReader(css));
        CSSOMParser parser = new CSSOMParser(new SACParserCSS3());

        CSSStyleSheet styleSheet1 = parser.parseStyleSheet(inputSource, null, null);

        CSSRuleList rules = styleSheet1.getCssRules();
        for (int i = 0; i < rules.getLength(); i++) {
            final CSSRule rule = rules.item(i);

            String ruleName = rule.getCssText().split("[{]")[0].strip().trim();
            cssRuleDeclaration.put(ruleName, new HashMap<>(12));

            String ruleDeclaration = rule.getCssText().replace("}", "");
            ruleDeclaration = ruleDeclaration.split("[{]")[1].trim().strip();

            InputSource source = new InputSource(new StringReader(ruleDeclaration));
            CSSStyleDeclaration declaration = parser.parseStyleDeclaration(source);

            for (int j = 0; j < declaration.getLength(); j++) {
                final String propName = declaration.item(j).trim().strip();
                String propertyValue = declaration.getPropertyValue(propName);
                cssRuleDeclaration.get(ruleName).put(propName, propertyValue.trim().strip());
            }
        }
        CellStyleProcessor cellStyleProcessor = new CellStyleProcessor();
        cellStyleProcessor.cssRuleDeclaration = cssRuleDeclaration;
        cellStyleProcessor.xssfWorkbook = xssfWorkbook;
        return cellStyleProcessor;
    }

    //-------------------------- Public Usage -------------------------------------------------

    private CellStyleProcessor() {
    }

    public XSSFCellStyle createStyle(String css) {
        css = css.trim().strip();

        if (cacheCellStyle.containsKey(css)) {
            return cacheCellStyle.get(css);
        }

        XSSFCellStyle xssfCellStyle = this.xssfWorkbook.createCellStyle();
        Map<String, String> cssProperties = this.cssRuleDeclaration.getOrDefault(css.trim().strip(), DEFAULT_CSS_PROPERTIES);

        createBorderFromCssInstructions(cssProperties, xssfCellStyle);
        createAlignementFromCssInstruction(cssProperties, xssfCellStyle);
        createCellBackgroundFromCssInstruction(cssProperties, xssfCellStyle);
        createFontFromCssInstruction(cssProperties, xssfCellStyle);

        cacheCellStyle.put(css, xssfCellStyle);
        return xssfCellStyle;
    }

    //-------------------------- Private Usage -------------------------------------------------
    private void createBorderFromCssInstructions(Map<String, String> instructions,
                                                 XSSFCellStyle xssfCellStyle) {
        instructions.forEach((prop, value) -> {
            if (prop.equals(BORDER)) {
                XSSFColor xssfColor = getBorderColor(value);
                xssfCellStyle.setBorderBottom(getBorderStyle(value));
                xssfCellStyle.setBottomBorderColor(xssfColor);

                xssfCellStyle.setBorderTop(getBorderStyle(value));
                xssfCellStyle.setTopBorderColor(xssfColor);

                xssfCellStyle.setBorderLeft(getBorderStyle(value));
                xssfCellStyle.setLeftBorderColor(xssfColor);

                xssfCellStyle.setBorderRight(getBorderStyle(value));
                xssfCellStyle.setRightBorderColor(xssfColor);

            } else if (prop.contains(BORDER_DASH)) {
                String side = prop.split(DASH)[1];
                XSSFColor xssfColor = getBorderColor(value);
                switch (side) {
                    case CssConstantes.LEFT -> {
                        xssfCellStyle.setBorderBottom(getBorderStyle(value));
                        xssfCellStyle.setBottomBorderColor(xssfColor);
                    }
                    case TOP -> {
                        xssfCellStyle.setBorderTop(getBorderStyle(value));
                        xssfCellStyle.setTopBorderColor(xssfColor);
                    }
                    case BOTTOM -> {
                        xssfCellStyle.setBorderLeft(getBorderStyle(value));
                        xssfCellStyle.setLeftBorderColor(xssfColor);
                    }
                    default -> {
                        xssfCellStyle.setBorderRight(getBorderStyle(value));
                        xssfCellStyle.setRightBorderColor(xssfColor);
                    }
                }
            }
        });
    }

    private BorderStyle getBorderStyle(String css) {
        String borderPropertiesWithoutColor = Arrays.stream(css.split("[ ]"))
                .filter(e -> e.contains("1p") || e.contains("2p") || e.contains("sol") || e.contains("das"))
                .collect(joining(" "));
        return CSS_BORDER_VALUE_TO_BORDER_STYLE_MAP.getOrDefault(borderPropertiesWithoutColor, NONE);

    }

    private XSSFColor getBorderColor(String css) {
        XSSFColor xssfColor = new XSSFColor();
        String[] borderStyleComposition = css.split(SPACE);
        String borderColor = Arrays.stream(borderStyleComposition)
                .filter(e -> e.contains(HASHBANG))
                .findFirst()
                .orElse(NORMAL_GRAY_BORDER_COLOR_HEX);
        xssfColor.setARGBHex(borderColor.replace(HASHBANG, EMPTY));
        return xssfColor;
    }

    private HorizontalAlignment getHorizontalAlignmentFromString(String textAlignment) {
        return switch (ofNullable(textAlignment).orElse(EMPTY)) {
            case TEXT_ALIGN_CENTER -> HorizontalAlignment.CENTER;
            case TEXT_ALIGN_RIGHT -> RIGHT;
            default -> LEFT;
        };
    }

    private XSSFColor getXSSFColorFromRgb(String rgb) {
        XSSFColor xssfColor = new XSSFColor();
        xssfColor.setARGBHex(HEX_WHITE);

        if (StringUtils.isEmpty(rgb)) {
            xssfColor.setARGBHex(HEX_WHITE);
        } else if (rgb.contains("rgb")) {
            rgb = rgb.replace("rgb(", "").replace(")", "");
            String[] rgbArr = rgb.split(",");

            int r = parseInt(rgbArr[0].trim().strip());
            int g = parseInt(rgbArr[1].trim().strip());
            int b = parseInt(rgbArr[2].trim().strip());

            Color color = new Color(r, g, b);
            String buf = Integer.toHexString(color.getRGB());
            String hex = buf.substring(buf.length() - 6);
            xssfColor.setARGBHex(hex);
        }
        return xssfColor;
    }

    private boolean isBold(Map<String, String> cssProperties) {
        return ofNullable(cssProperties.get(FONT_WEIGHT))
                .map(fw -> fw.contains(BOLD))
                .orElse(false);
    }

    private void createAlignementFromCssInstruction(Map<String, String> instructions,
                                                    XSSFCellStyle xssfCellStyle) {
        String textAlignment = instructions.get(TEXT_ALIGN);
        xssfCellStyle.setVerticalAlignment(CENTER);
        xssfCellStyle.setAlignment(getHorizontalAlignmentFromString(textAlignment));
    }

    private void createCellBackgroundFromCssInstruction(Map<String, String> instructions,
                                                        XSSFCellStyle xssfCellStyle) {
        String cellBackground = instructions.get(BACKGROUND);
        xssfCellStyle.setFillForegroundColor(getXSSFColorFromRgb(cellBackground));
        xssfCellStyle.setFillPattern(SOLID_FOREGROUND);
    }

    private void createFontFromCssInstruction(Map<String, String> instructions,
                                              XSSFCellStyle xssfCellStyle) {
        XSSFColor textColor = getXSSFColorFromRgb(instructions.get(COLOR));
        boolean isBold = isBold(instructions);

        XSSFFont xssfFont = xssfWorkbook.createFont();
        xssfFont.setColor(textColor);
        xssfFont.setBold(isBold);
        xssfCellStyle.setFont(xssfFont);
    }
}