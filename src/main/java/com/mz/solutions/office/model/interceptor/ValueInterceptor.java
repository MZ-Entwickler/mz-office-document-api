/*
 * MZ Office Document API
 *
 * Moritz Riebe und Andreas Zaschka GbR
 *
 * Copyright (C) 2020,   Moritz Riebe     (moritz.riebe@mz-entwickler.de)
 *                       Andreas Zaschka  (andreas.zaschka@mz-entwickler.de)
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Affero General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Affero General Public License for more details.
 * 
 * You should have received a copy of the GNU Affero General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
package com.mz.solutions.office.model.interceptor;

import com.mz.solutions.office.extension.ExtendedValue;
import com.mz.solutions.office.model.DataValue;
import com.mz.solutions.office.model.DataValueMap;
import java.util.Objects;

/**
 * Bei Bedarf spät (lazy) evaluierte Platzhalter im Ersetzungsvorgang für aufwendigere Ersetzungen
 * die nur bei Bedarf ausgeführt werden sollen oder Kontaxt abhängige Vorgänge.
 * 
 * <p>Entspricht einem {@link ExtendedValue} und wird/ muss von allen Office Implementierungen
 * unterstützt werden. Wird in {@link DataValue} mit einem {@link ValueInterceptor} erzeugt, so
 * erfolgt keine Zwischenspeicherung (caching), aber eine Generierung des Platzhalterwertes nur
 * bei Bedarf (lazy-evaluation). Der "Ergebnisplatzhalter" (ein {@link DataValue} das verwendet
 * werden soll), kann ebenfalls ein Interceptor sein - entsprechende Vorgänge werden so lange
 * aufgelöst bis ein direkt verwendbarer {@link DataValue} zurückgegeben wird.</p>
 * 
 * <p>Zur Generierung/ Evaluierung des eigentlich einzusetzenden Wertes wird
 * {@link #interceptPlaceholder(InterceptionContext)} aufgerufen mit der Unterstützung durch einen
 * {@link InterceptionContext} für den Zugriff auf den eigentlichen Namen des Platzhalters, den
 * Scope im Ersetzungsvorgang {@link DataValueMap} und erweiterten Funktionen.</p>
 * 
 * <p><b>Beispiel:</b> Zwei Platzhalter die normal sind, ein Platzhalter der nur bei Bedarf 
 * (also gefundenem Platzhalter im Dokument) erzeugt wird:</p>
 * 
 * <pre>
 *  final OfficeDocumentFactory docFactory = ...
 *  final OfficeDocument document = docFactory. ..
 * 
 *  final DataPage page = new DataPage();
 *  page.addValue(new DataValue("FORENAME", "John"));
 *  page.addValue(new DataValue("LASTNAME", "Doe"));
 * 
 *  // lazy evaluated value if this placeholder will be used
 *  // Value-Interceptor using Lambda's
 *  page.addValue(new DataValue("FULLNAME", ValueInterceptor.callFunction(innerContext -&gt; {
 *      final InterceptionContext context = innerContext;
 *      final String nameOfPlaceholder = context.getPlaceholderName();
 *      final DataValueMap&lt;?&gt; valueMap = context.getParentValueMap(); // is "DataPage page" !
 * 
 *      final String forename = valueMap.getValueByKey("FORENAME").get().getValue();
 *      final String lastname = valueMap.getValueByKey("LASTNAME").get().getValue();
 * 
 *      final String fullname = lastname + ", " + forename;
 * 
 *      return DataValueResult.useValue(new DataValue(nameOfPlaceholder, fullname));
 *  })));
 * 
 *  // ...
 * 
 *  // Lambdas are possible, but you can also directly inherit abstract class ValueInterceptor
 *  
 *  class FullNameInterceptor extends ValueInterceptor {
 *
 *      public DataValueResult interceptPlaceholder(InterceptionContext context) {
 *          final InterceptionContext context = innerContext;
 *          final String nameOfPlaceholder = context.getPlaceholderName();
 *          final DataValueMap&lt;?&gt; valueMap = context.getParentValueMap(); // is "DataPage page" !
 *     
 *          final String forename = valueMap.getValueByKey("FORENAME").get().getValue();
 *          final String lastname = valueMap.getValueByKey("LASTNAME").get().getValue();
 *     
 *          final String fullname = lastname + ", " + forename;
 *     
 *          return DataValueResult.useValue(new DataValue(nameOfPlaceholder, fullname));
 *      }
 * 
 *  }
 * 
 *  page.addValue(new DataValue("FULLNAME", new FullNameInterceptor()));
 * </pre>
 * 
 * @see InterceptionContext         Kontext während des Evaluierungsvorganges
 * @see ValueInterceptorFunction    Methoden-Signatur für Interceptors
 * 
 * @author Riebe, Moritz (moritz.riebe@mz-entwickler.de)
 */
public abstract class ValueInterceptor extends ExtendedValue implements ValueInterceptorFunction {
    
    /**
     * Erzeugt einen Value-Interceptor unter Verwendung Lamdas.
     * 
     * @param altText               Alternativer Text im Falle fehlender Unterstützung,
     *                              sollte nie {@code null} sein.
     * 
     * @param interceptorFunction   Funktion für die Erzeugung des Platzhalterwertes.
     * 
     * @return                      Instanz von {@link ValueInterceptor} welche die übergebene
     *                              Funktion zur Werterzuegung aufruft.
     */
    public static ValueInterceptor callFunction(
            final String altText,
            final ValueInterceptorFunction interceptorFunction)
    {
        return new FunctionBasedValueInterceptor(interceptorFunction, altText);
    }
    
    /**
     * Erzeugt einen Value-Interceptor unter Verwendung von Lambdas und ohne alternativen
     * Text.
     * 
     * <p>Verhält sich identisch zu {@link #callFunction(String, ValueInterceptorFunction)} unter
     * Angabe von {@code altText = ""}.</p>
     * 
     * @param interceptorFunction   Funktion für die Erzeugung des Platzhalterwertes.
     * 
     * @return                      Instanz von {@link ValueInterceptor} welche die übergebene
     *                              Funktion zur Werterzuegung aufruft.
     */
    public static ValueInterceptor callFunction(
            final ValueInterceptorFunction interceptorFunction)
    {
        return new FunctionBasedValueInterceptor(interceptorFunction);
    }
    
    private static class FunctionBasedValueInterceptor extends ValueInterceptor {
        
        private final ValueInterceptorFunction function;

        public FunctionBasedValueInterceptor(
                final ValueInterceptorFunction function,
                final String altText)
        {
            super(altText);
            this.function = Objects.requireNonNull(function, "function");
        }

        public FunctionBasedValueInterceptor(
                final ValueInterceptorFunction function)
        {
            this(function, "");
        }

        @Override
        public DataValueResult interceptPlaceholder(InterceptionContext context) {
            return function.interceptPlaceholder(context);
        }
        
    }
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    
    private final String altText;
    
    /**
     * Value-Interceptor mit alternativem Text für Office-Implementierungen ohne eine solche
     * Unterstützung.
     * 
     * @param altText       Alternativer Text, nicht {@code null}, aber
     *                      {@code string.isEmpty() == true} ist möglich.
     */
    public ValueInterceptor(String altText) {
        this.altText = (null == altText) ? "" : altText;
    }
    
    /**
     * Value-Interceptor ohne alternativen Text.
     */
    public ValueInterceptor() {
        this(null);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String altString() {
        return altText;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public abstract DataValueResult interceptPlaceholder(InterceptionContext context)
            throws ValueInterceptorException;
    
}
